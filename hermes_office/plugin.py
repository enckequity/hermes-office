from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from . import bridge, settings

TOOLSET = "office"
SKILL_DIR = Path(__file__).parent / "skills" / "office-review"

logger = logging.getLogger(__name__)


def engines_ready() -> bool:
    """True when the Paper forks (not stock python-docx/pptx/openpyxl) are the importable packages."""
    try:
        import docx
        import openpyxl
        import pptx
    except Exception:  # absent, or the Paper forks' own refusal to share a package with the stock library
        return False
    return all(getattr(module, "__paper_version__", None) for module in (docx, pptx, openpyxl))


def _desk_exists() -> bool:
    """A Desk exists once a folder or edit created it; per-turn hooks never create one."""
    try:
        return (bridge.data_dir() / "desk.db").exists()
    except Exception:
        return False


def _on_turn_start(session_id: str = "", user_message: Any = None, **_: Any) -> None:
    if not session_id:
        return
    from . import service

    service.turn_started(session_id, user_message if isinstance(user_message, str) else "")
    if not _desk_exists():
        return
    try:
        d = service.desk()
        ws = d.workstream_for_session(session_id)
        if ws and not ws.get("needs_you") and ws.get("status") != "running":
            service.set_status(d, ws["id"], "running")
    except Exception:  # Desk bookkeeping must never disturb a turn
        logger.debug("office-desk: turn-start bookkeeping failed", exc_info=True)


def _on_turn_end(session_id: str = "", **_: Any) -> None:
    from . import service

    if not session_id or service.turn_ended(session_id) is None or not _desk_exists():
        return
    try:
        d = service.desk()
        ws = d.workstream_for_session(session_id)
        if not ws or ws.get("needs_you"):
            return
        last = ws.get("last_event") or {}
        if last.get("kind") == "refused":
            service.set_status(d, ws["id"], "needs_you", note=f"An edit was refused: {last.get('summary', '')}"[:300], needs_you=True)
        else:
            service.set_status(d, ws["id"], "done")
    except Exception:
        logger.debug("office-desk: turn-end bookkeeping failed", exc_info=True)


def register(ctx) -> None:
    settings.bind(ctx)
    if not engines_ready():
        # The tool modules import Paper-only APIs (docx.commentops, openpyxl.oracle, pptx.diff, ...), so with the
        # stock libraries they cannot even be imported. Register nothing: no tools, no skill, no per-turn hooks.
        logger.warning("office-desk: the Paper Office engines (paper-docx, paper-pptx, paper-xlsx) are not the importable "
                       "docx/pptx/openpyxl packages here; the plugin registers nothing")
        return
    from .schemas import ALL_SCHEMAS
    from .tools import HANDLERS

    for name, handler in HANDLERS.items():
        ctx.register_tool(name=name, toolset=TOOLSET, schema=ALL_SCHEMAS[name], handler=handler, check_fn=engines_ready, emoji="🗂️")
    ctx.register_skill("office-review", SKILL_DIR / "SKILL.md", description="Review and edit Word, PowerPoint and Excel files on the Desk with tracked changes and receipts.")
    ctx.register_hook("pre_llm_call", _on_turn_start)
    ctx.register_hook("post_llm_call", _on_turn_end)
