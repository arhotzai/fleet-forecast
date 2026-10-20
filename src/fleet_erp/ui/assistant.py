"""
In-app Claude assistant — ask in plain language for the data you want, in the
structure you want. Claude writes a pandas query against the projection results,
it runs in a guarded namespace, and the answer + resulting table come back.

Uses the official Anthropic SDK (Claude Opus 4.8) with a manual tool-use loop.
"""

from __future__ import annotations

import re

import pandas as pd

MODEL = "claude-opus-4-8"

# Blocked substrings in generated code (defense-in-depth for a local tool).
_FORBIDDEN = re.compile(
    r"(__|\bimport\b|\bopen\b|\beval\b|\bexec\b|\bos\b|\bsys\b|"
    r"\bsubprocess\b|\bgetattr\b|\bsetattr\b|\bglobals\b|\blocals\b|"
    r"\bcompile\b|\binput\b|\bfile\b)"
)

_QUERY_TOOL = {
    "name": "query_data",
    "description": (
        "Run a short pandas snippet against the fleet projection DataFrames and "
        "return the result. The snippet MUST assign the final answer to a variable "
        "named `result` (a DataFrame or Series). Available names: pd (pandas) and the "
        "DataFrames listed in the system prompt. No imports, file, or system access."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "code": {"type": "string", "description": "pandas snippet assigning to `result`"},
            "explanation": {"type": "string", "description": "one line on what this computes"},
        },
        "required": ["code"],
    },
}


def available_frames(ds, proj) -> dict[str, pd.DataFrame]:
    """The DataFrames the assistant can query."""
    frames = {
        "events": proj.events if proj.events is not None else pd.DataFrame(),
        "sv_calendar": proj.sv_plan_v1 if proj.sv_plan_v1 is not None else pd.DataFrame(),
        "sv_matrix": proj.sv_plan_v2 if proj.sv_plan_v2 is not None else pd.DataFrame(),
        "unit_status": proj.unit_status if proj.unit_status is not None else pd.DataFrame(),
        "fleet": ds.get("Fleet"),
        "components": ds.get("Components"),
        "llp": ds.get("LLP_Registry"),
        "utilization": ds.get("Utilization"),
    }
    return {k: v for k, v in frames.items() if v is not None}


def _schema_text(frames: dict[str, pd.DataFrame]) -> str:
    lines = []
    for name, df in frames.items():
        cols = ", ".join(map(str, df.columns))[:400]
        lines.append(f"- `{name}` ({len(df)} rows): {cols}")
    return "\n".join(lines)


def _system_prompt(frames) -> str:
    return (
        "You are a data assistant embedded in a Fleet Maintenance ERP. The user asks "
        "in plain language for maintenance-projection data in whatever structure they "
        "want. Use the `query_data` tool to compute the answer with pandas, then give a "
        "short natural-language summary. Prefer returning a tidy DataFrame as `result`.\n\n"
        "Key facts: each row in `events` is a projected shop visit or LLP replacement "
        "(columns Date, Tail, Position [LHENG/RHENG/APU/AIRFRAME], Unit_ID, Unit_Type, "
        "Event, Driver, SLSV_at_SV, Trigger [LIMIT/HARD_LINK/SOFT_BUNDLE/LLP_REPLACEMENT]). "
        "`llp` has one row per life-limited part with ToGo (cycles remaining) and SLSV. "
        "Dates are pandas Timestamps.\n\n"
        "Available DataFrames:\n" + _schema_text(frames) +
        "\n\nAlways assign the final table to `result`. Keep snippets to a few lines. "
        "No imports or file access — `pd` and the DataFrames are already in scope."
    )


def run_query(code: str, frames: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Execute a generated pandas snippet in a guarded namespace; return `result`."""
    if _FORBIDDEN.search(code):
        raise ValueError("Generated code used a disallowed operation (imports/file/system access).")
    ns = {"pd": pd, **{k: v.copy() for k, v in frames.items()}}
    safe_builtins = {"len": len, "range": range, "min": min, "max": max, "sum": sum,
                     "sorted": sorted, "list": list, "dict": dict, "set": set,
                     "round": round, "abs": abs, "str": str, "int": int, "float": float,
                     "bool": bool, "zip": zip, "enumerate": enumerate, "any": any, "all": all}
    exec(code, {"__builtins__": safe_builtins}, ns)   # noqa: S102 — guarded, local single-user tool
    result = ns.get("result")
    if result is None:
        raise ValueError("The snippet did not assign a `result` variable.")
    if isinstance(result, pd.Series):
        result = result.to_frame()
    if not isinstance(result, pd.DataFrame):
        result = pd.DataFrame({"result": [result]})
    return result


def ask(client, frames: dict[str, pd.DataFrame], history: list[dict]):
    """
    Run one assistant turn with a manual tool-use loop.
    `history` is the running list of Anthropic message dicts (mutated in place).
    Returns (answer_text, result_df_or_None, last_code_or_None).
    """
    system = _system_prompt(frames)
    result_df = None
    last_code = None

    for _ in range(5):  # cap tool round-trips
        resp = client.messages.create(
            model=MODEL,
            max_tokens=4000,
            thinking={"type": "adaptive"},
            system=system,
            tools=[_QUERY_TOOL],
            messages=history,
        )
        history.append({"role": "assistant", "content": resp.content})

        if resp.stop_reason != "tool_use":
            text = "".join(b.text for b in resp.content if b.type == "text")
            return text, result_df, last_code

        tool_results = []
        for block in resp.content:
            if block.type != "tool_use":
                continue
            code = block.input.get("code", "")
            last_code = code
            try:
                result_df = run_query(code, frames)
                preview = result_df.head(30).to_csv(index=False)
                content = f"Query succeeded. First rows (CSV):\n{preview}\n(total rows: {len(result_df)})"
                is_err = False
            except Exception as exc:  # feed the error back so Claude can fix it
                content = f"Error running the snippet: {exc}"
                is_err = True
            tool_results.append({
                "type": "tool_result", "tool_use_id": block.id,
                "content": content, "is_error": is_err,
            })
        history.append({"role": "user", "content": tool_results})

    return "I couldn't complete that after several attempts — try rephrasing.", result_df, last_code
