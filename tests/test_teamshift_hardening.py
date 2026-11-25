"""TeamShift hardening (enckequity/teamshift#4657): no agent-widened Desk, neutral author and recalc text,
nothing registered without the Paper engines, per-turn hooks that never create a Desk, no engine name in user text."""

from __future__ import annotations

import ast
import json
import subprocess
import sys
from pathlib import Path
from typing import Any

import pytest

from conftest import REPO, call
from hermes_office import bridge, plugin, settings, xlsx_ops
from hermes_office.guard import sha256
from hermes_office.schemas import ALL_SCHEMAS

EXPECTED_TOOLS = {
    "office_desk", "office_inspect", "office_search", "docx_tracked_edit", "docx_comment", "docx_redline_accept_reject",
    "pptx_inspect_formats", "pptx_edit_shape", "pptx_compose_decks", "pptx_diff", "xlsx_edit", "xlsx_recalc", "xlsx_diff",
    "office_receipt",
}


class RecordingCtx:
    def __init__(self, values: dict[str, Any] | None = None):
        self.values = values or {}
        self.tools: list[dict[str, Any]] = []
        self.hooks: list[str] = []
        self.skills: list[str] = []

    def get_config(self, key: str, default: Any = None) -> Any:
        return self.values.get(key, default)

    def register_tool(self, **kwargs: Any) -> None:
        self.tools.append(kwargs)

    def register_hook(self, name: str, callback: Any) -> None:
        self.hooks.append(name)

    def register_skill(self, name: str, path: Path, **_: Any) -> None:
        self.skills.append(name)


@pytest.fixture(autouse=True)
def _clean_settings():
    settings._reset_for_tests()
    yield
    settings._reset_for_tests()


# add_folder is gone -------------------------------------------------------------------------------------------------

def test_office_desk_schema_has_no_add_folder() -> None:
    schema = ALL_SCHEMAS["office_desk"]
    assert "add_folder" not in schema["parameters"]["properties"]["action"]["enum"]
    assert "path" not in schema["parameters"]["properties"]
    assert "add_folder" not in json.dumps(ALL_SCHEMAS)


def test_add_folder_action_is_refused_and_desk_unchanged(files, desk, tmp_path: Path) -> None:
    other = tmp_path / "Somewhere Else"
    other.mkdir()
    before = desk.list_folders()
    out = call("office_desk", {"action": "add_folder", "path": str(other)})
    assert out["ok"] is False
    assert desk.list_folders() == before
    target = other / "x.docx"
    target.write_bytes((files["docx"]).read_bytes())
    refused = call("docx_tracked_edit", {"path": str(target), "find": "within 30 days", "replace": "within 45 days"})
    assert refused["ok"] is False and refused["reason"] == "outside_desk"
    assert "add_folder" not in refused.get("hint", "")
    assert sha256(target) == sha256(files["docx"])


# default author ------------------------------------------------------------------------------------------------------

@pytest.mark.parametrize("configured, expected", [
    (None, "Assistant"),
    ("", "Assistant"),
    ("   ", "Assistant"),
    ("Hermes", "Assistant"),
    ("hermes", "Assistant"),
    ("Northwind Plumbing", "Northwind Plumbing"),
])
def test_default_author(hermes_home, configured, expected) -> None:
    settings.bind(RecordingCtx({} if configured is None else {"default_author": configured}))
    assert settings.default_author() == expected


def test_tracked_edit_uses_business_name_and_never_hermes(files) -> None:
    import docx

    call("docx_tracked_edit", {"path": str(files["docx"]), "find": "within 30 days", "replace": "within 45 days"})
    authors = {r.author for r in docx.Document(str(files["docx"])).revisions}
    assert authors == {"Assistant"}
    settings.bind(RecordingCtx({"default_author": "Northwind Plumbing"}))
    call("docx_comment", {"path": str(files["docx"]), "find": "Schedule B", "text": "Check this"})
    out = call("office_inspect", {"path": str(files["docx"]), "detail": "full"})
    assert "Hermes" not in json.dumps(out)
    assert "Northwind Plumbing" in json.dumps(out)


# recalc without LibreOffice ---------------------------------------------------------------------------------------------

def test_recalc_without_soffice_is_neutral_and_changes_nothing(files, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(xlsx_ops, "soffice_path", lambda: None)
    path = files["xlsx"]
    before = sha256(path)
    out = call("xlsx_recalc", {"path": str(path), "read": ["Budget!F2"]})
    assert out["ok"] is True and out["recalculated"] is False
    assert out["reason"] == "recalc_unavailable"
    assert "LibreOffice" not in out["hint"] and "soffice" not in out["hint"] and "Install" not in out["hint"]
    assert "Budget!F2" in out["values"]
    assert sha256(path) == before
    assert out["changed"] is False


# registration ----------------------------------------------------------------------------------------------------------

def test_register_counts_with_engines(hermes_home) -> None:
    ctx = RecordingCtx()
    plugin.register(ctx)
    assert {t["name"] for t in ctx.tools} == EXPECTED_TOOLS
    assert len(ctx.tools) == 14
    assert {t["toolset"] for t in ctx.tools} == {"office"}
    assert ctx.hooks == ["pre_llm_call", "post_llm_call"]
    assert ctx.skills == ["office-review"]


def test_register_nothing_without_engines(hermes_home, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(plugin, "engines_ready", lambda: False)
    ctx = RecordingCtx()
    plugin.register(ctx)
    assert ctx.tools == [] and ctx.hooks == [] and ctx.skills == []


STUB_PROBE = """
import json, sys
sys.path.insert(0, sys.argv[1])
sys.path.insert(0, sys.argv[2])
import plugin_pkg
class Ctx:
    def __init__(self): self.n = {"tools": 0, "hooks": 0, "skills": 0}
    def get_config(self, key, default=None): return default
    def register_tool(self, **kw): self.n["tools"] += 1
    def register_hook(self, *a, **kw): self.n["hooks"] += 1
    def register_skill(self, *a, **kw): self.n["skills"] += 1
ctx = Ctx()
plugin_pkg.register(ctx)
loaded = sorted(m for m in sys.modules if m.split(".")[0] in ("docx", "pptx", "openpyxl") or "dashboard" in m or "plugin_api" in m)
print("RESULT:" + json.dumps({"counts": ctx.n, "loaded": loaded}))
"""


def test_stock_libraries_load_cleanly_and_register_nothing(tmp_path: Path) -> None:
    """With stock python-docx/openpyxl/python-pptx (no __paper_version__), loading the plugin must not crash."""
    stubs = tmp_path / "stubs"
    for name in ("docx", "openpyxl", "pptx"):
        (stubs / name).mkdir(parents=True)
        (stubs / name / "__init__.py").write_text("__version__ = 'stock'\n")
    host = tmp_path / "host"
    host.mkdir()
    (host / "plugin_pkg").symlink_to(REPO, target_is_directory=True)
    probe = tmp_path / "probe.py"
    probe.write_text(STUB_PROBE)
    result = subprocess.run([sys.executable, "-I", str(probe), str(stubs), str(host)], capture_output=True, text=True, timeout=60,
                            env={"HERMES_HOME": str(tmp_path / "home"), "PATH": "/usr/bin:/bin"})
    assert result.returncode == 0, result.stderr[-3000:]
    line = next(line for line in result.stdout.splitlines() if line.startswith("RESULT:"))
    report = json.loads(line.removeprefix("RESULT:"))
    assert report["counts"] == {"tools": 0, "hooks": 0, "skills": 0}, report
    assert all(m in ("docx", "pptx", "openpyxl") for m in report["loaded"]), report  # only the readiness probe
    assert not (tmp_path / "home").exists(), "loading must not create plugin data"


# hooks ------------------------------------------------------------------------------------------------------------------

def test_turn_hooks_never_create_a_desk(hermes_home) -> None:
    db = bridge.data_dir() / "desk.db"
    plugin._on_turn_start(session_id="chat-x", user_message="hello")
    plugin._on_turn_end(session_id="chat-x")
    assert not db.exists() and not db.parent.exists()


def test_turn_hooks_swallow_desk_errors(desk, monkeypatch: pytest.MonkeyPatch) -> None:
    from hermes_office import service

    def boom() -> None:
        raise RuntimeError("db locked")

    monkeypatch.setattr(service, "desk", boom)
    plugin._on_turn_start(session_id="chat-y", user_message="hello")
    plugin._on_turn_end(session_id="chat-y")


# runtime does not import the stripped dashboard / desktop ------------------------------------------------------------------

def test_runtime_modules_do_not_import_dashboard_or_desktop() -> None:
    runtime = [REPO / "__init__.py", *sorted((REPO / "hermes_office").glob("*.py"))]
    for path in runtime:
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            names = []
            if isinstance(node, ast.Import):
                names = [a.name for a in node.names]
            elif isinstance(node, ast.ImportFrom):
                names = [node.module or ""]
            for name in names:
                assert not name.split(".")[0] in ("dashboard", "desktop", "plugin_api", "fastapi"), (path.name, name)


# white-label --------------------------------------------------------------------------------------------------------------

def _strings(value: Any):
    if isinstance(value, str):
        yield value
    elif isinstance(value, dict):
        for v in value.values():
            yield from _strings(v)
    elif isinstance(value, list):
        for v in value:
            yield from _strings(v)


def test_user_and_model_facing_text_does_not_name_hermes() -> None:
    from hermes_office import errors, guard

    texts = list(_strings(ALL_SCHEMAS))
    texts.append((REPO / "hermes_office" / "skills" / "office-review" / "SKILL.md").read_text(encoding="utf-8"))
    texts += [cls.hint for cls in vars(errors).values() if isinstance(cls, type) and hasattr(cls, "hint")]
    from openpyxl.errors import OracleUnavailableError

    texts.append(guard._hint_for_exception(OracleUnavailableError("no soffice")))
    try:
        import yaml
        manifest = yaml.safe_load((REPO / "plugin.yaml").read_text(encoding="utf-8"))
    except ImportError:
        from ruamel.yaml import YAML
        manifest = YAML(typ="safe").load((REPO / "plugin.yaml").read_text(encoding="utf-8"))
    texts += list(_strings(manifest["config_schema"]))
    for text in texts:
        assert "Hermes" not in text, text
