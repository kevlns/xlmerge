"""Deterministic validation of v-cli.plugin.json against the argparse CLI and package metadata.

Covers:
- manifest shape per the v-cli contract (schemaVersion 1), including Commander-style
  display-string flags (e.g. "--repo <path>", "--no-browser")
- package pointer (package.json vCli + files) and bin consistency
- two-way drift check: every public CLI subcommand/option is in the manifest and
  every manifest command/option exists in the parser
- side-effect markers match the real behavior of each command
- regression: --repo defaults to the current directory (previously a literal comma)
"""

import argparse
import json
import re
import sys
import unittest
from pathlib import Path

PACKAGE_ROOT = Path(__file__).resolve().parents[2]
RESOLVER_DIR = PACKAGE_ROOT / "python" / "xlsx_resolver"
sys.path.insert(0, str(RESOLVER_DIR))

from resolve_xlsx_conflict import build_parser

PUBLIC_COMMANDS = ["detect", "prepare", "resolve", "launch", "apply"]
HIDDEN_COMMANDS = {"serve"}  # internal subcommand spawned by launch (help=SUPPRESS)

TOP_LEVEL_KEYS = ["schemaVersion", "package", "command", "bin", "description", "platforms", "runtime", "environment", "agent"]
AGENT_KEYS = ["whenToUse", "globalOptions", "commands"]
COMMAND_KEYS = ["path", "usage", "description", "arguments", "options", "output", "exitCodes", "safety"]
OPTION_KEYS = ["flags", "description"]

# documented effect tokens; the manifest must only use these
SAFETY_TOKENS = {
    "read-only",
    "no-worktree-modification",
    "writes-runtime-dir",
    "no-commit",
    "blocking",
    "binds-loopback",
    "opens-browser-by-default",
    "no-push",
    "starts-background-server",
    "writes-worktree",
    "writes-worktree-via-ui",
    "commits-by-default",
    "commits-by-default-via-ui",
    "pushes-only-with-flag",
}


def public_subparsers(parser):
    """Return {name: subparser} for every non-hidden subcommand."""
    result = {}
    for action in parser._actions:
        if not isinstance(action, argparse._SubParsersAction):
            continue
        hidden = {item.dest for item in action._choices_actions if item.help == argparse.SUPPRESS}
        for name, subparser in action.choices.items():
            if name not in hidden and name not in ("help", "?"):
                result[name] = subparser
    return result


def parser_option_keys(subparser):
    """Option-string tuples of a subparser, excluding the auto -h/--help action."""
    keys = set()
    for action in subparser._actions:
        if action.option_strings and tuple(action.option_strings) != ("-h", "--help"):
            keys.add(tuple(action.option_strings))
    return keys


def manifest_flag_keys(options):
    """Flag tuples parsed from Commander-style display-string entries; "<meta>" placeholders are dropped."""
    keys = set()
    for option in options:
        keys.add(tuple(token for token in option["flags"].split() if not token.startswith("<")))
    return keys


class VCliPluginManifestTests(unittest.TestCase):
    def setUp(self):
        self.manifest = json.loads((PACKAGE_ROOT / "v-cli.plugin.json").read_text(encoding="utf-8"))
        self.package = json.loads((PACKAGE_ROOT / "package.json").read_text(encoding="utf-8"))

    def _assert_flags_display_string(self, flags):
        """flags must be a deterministic Commander-style display string, not an array."""
        self.assertIsInstance(flags, str)
        self.assertEqual(flags.strip(), flags)
        self.assertNotIn("  ", flags)
        self.assertTrue(flags.split()[0].startswith("-"))

    def test_manifest_matches_package_metadata_and_shipping_rules(self):
        manifest, package = self.manifest, self.package
        self.assertEqual(manifest["package"], package["name"])
        self.assertEqual(package["name"], "@kevlns/xlmerge")
        self.assertEqual(manifest["command"], "xlmerge")
        self.assertEqual(manifest["bin"], "xlmerge")
        self.assertEqual(package["bin"], {"xlmerge": "bin/xlmerge.js"})
        self.assertTrue((PACKAGE_ROOT / "bin" / "xlmerge.js").is_file())
        self.assertEqual(package["vCli"], {"manifest": "v-cli.plugin.json"})
        for required in ("v-cli.plugin.json", "AGENTS.md", "bin/", "python/xlsx_merge_engine/", "python/xlsx_resolver/"):
            self.assertIn(required, package["files"])
        for excluded in ("skill/", "python/tests", "e2e_test.py", "python/tests/"):
            self.assertNotIn(excluded, package["files"])
        agent_doc = PACKAGE_ROOT / "AGENTS.md"
        self.assertTrue(agent_doc.is_file())
        self.assertTrue(agent_doc.read_text(encoding="utf-8").strip())
        # manifest is actually reachable at the pointer path
        self.assertTrue((PACKAGE_ROOT / package["vCli"]["manifest"]).is_file())
        # environment variables must be the ones the Node shim actually reads
        shim = (PACKAGE_ROOT / "bin" / "xlmerge.js").read_text(encoding="utf-8")
        for var in ("XLMERGE_PYTHON", "XLMERGE_VENV"):
            self.assertIn(var, shim)
        # version is a publishable prerelease semver (exact bump is a release step)
        self.assertRegex(package["version"], r"^\d+\.\d+\.\d+(-[0-9A-Za-z.-]+)?$")

    def test_schema_shape_is_deterministic(self):
        manifest = self.manifest
        self.assertEqual(manifest["schemaVersion"], 1)
        self.assertEqual(list(manifest.keys()), TOP_LEVEL_KEYS)
        self.assertTrue(manifest["description"])
        self.assertEqual(sorted(manifest["platforms"]), ["darwin", "linux", "win32"])
        self.assertEqual(sorted(manifest["runtime"].keys()), ["node", "python", "pythonDependencies"])
        self.assertIsInstance(manifest["runtime"]["pythonDependencies"], list)
        self.assertTrue(manifest["runtime"]["pythonDependencies"])

        env_entries = manifest["environment"]
        self.assertEqual([sorted(entry.keys()) for entry in env_entries], [["description", "name"]] * len(env_entries))
        self.assertEqual({entry["name"] for entry in env_entries}, {"XLMERGE_PYTHON", "XLMERGE_VENV"})
        self.assertTrue(all(entry["description"] for entry in env_entries))

        agent = manifest["agent"]
        self.assertEqual(list(agent.keys()), AGENT_KEYS)
        self.assertTrue(agent["whenToUse"])
        self.assertEqual([sorted(option.keys()) for option in agent["globalOptions"]], [["description", "flags"]] * len(agent["globalOptions"]))
        for option in agent["globalOptions"]:
            self._assert_flags_display_string(option["flags"])
        self.assertEqual(manifest_flag_keys(agent["globalOptions"]), {("--repo",), ("--runtime-dir",)})

        commands = agent["commands"]
        self.assertEqual([tuple(command["path"]) for command in commands], [(name,) for name in PUBLIC_COMMANDS])
        usages = [command["usage"] for command in commands]
        self.assertEqual(len(usages), len(set(usages)))
        for command in commands:
            self.assertEqual(list(command.keys()), COMMAND_KEYS)
            self.assertEqual(command["arguments"], [])
            self.assertTrue(command["usage"].startswith("xlmerge "))
            self.assertTrue(command["description"])
            self.assertEqual([sorted(option.keys()) for option in command["options"]], [["description", "flags"]] * len(command["options"]))
            for option in command["options"]:
                self._assert_flags_display_string(option["flags"])
            self.assertEqual(sorted(command["output"].keys()), ["description", "format"])
            self.assertTrue(command["output"]["description"])
            self.assertIn("0", command["exitCodes"])
            self.assertIn("1", command["exitCodes"])
            self.assertTrue(command["safety"])
            for token in command["safety"]:
                self.assertIn(token, SAFETY_TOKENS)

    def test_manifest_covers_every_public_command_and_option(self):
        parser = build_parser()
        public = public_subparsers(parser)
        self.assertEqual(set(public), set(PUBLIC_COMMANDS))
        # serve must stay hidden; if it becomes public the manifest must cover it
        for action in parser._actions:
            if isinstance(action, argparse._SubParsersAction):
                hidden = {item.dest: item.help for item in action._choices_actions}
                self.assertEqual(hidden.get("serve"), argparse.SUPPRESS)

        manifest_commands = {tuple(command["path"])[0]: command for command in self.manifest["agent"]["commands"]}
        self.assertEqual(set(manifest_commands), set(public))
        for name, subparser in public.items():
            self.assertEqual(
                manifest_flag_keys(manifest_commands[name]["options"]),
                parser_option_keys(subparser),
                f"option drift for command '{name}'",
            )

        # global option drift check
        global_keys = set()
        for action in parser._actions:
            if action.option_strings and tuple(action.option_strings) != ("-h", "--help"):
                global_keys.add(tuple(action.option_strings))
        self.assertEqual(global_keys, {("--repo",), ("--runtime-dir",)})
        self.assertEqual(
            manifest_flag_keys(self.manifest["agent"]["globalOptions"]),
            {("--repo",), ("--runtime-dir",)},
        )

    def test_side_effect_markers_are_accurate(self):
        commands = {tuple(command["path"])[0]: command for command in self.manifest["agent"]["commands"]}

        def safety(name):
            return set(commands[name]["safety"])

        self.assertLessEqual({"read-only", "no-worktree-modification"}, safety("detect"))
        self.assertFalse(safety("detect") & {"writes-worktree", "writes-runtime-dir", "commits-by-default", "writes-worktree-via-ui", "commits-by-default-via-ui"})
        self.assertLessEqual({"writes-runtime-dir", "no-worktree-modification", "no-commit"}, safety("prepare"))
        self.assertLessEqual({"blocking", "binds-loopback", "opens-browser-by-default", "no-push", "writes-runtime-dir", "writes-worktree-via-ui", "commits-by-default-via-ui"}, safety("resolve"))
        self.assertLessEqual({"starts-background-server", "binds-loopback", "opens-browser-by-default", "no-push", "writes-runtime-dir", "writes-worktree-via-ui", "commits-by-default-via-ui"}, safety("launch"))
        self.assertLessEqual({"writes-worktree", "commits-by-default", "pushes-only-with-flag"}, safety("apply"))
        self.assertFalse(safety("apply") & {"read-only", "no-commit", "no-push", "writes-runtime-dir"})

        # output formats match the CLI contract
        self.assertEqual(commands["detect"]["output"]["format"], "json")
        self.assertEqual(commands["prepare"]["output"]["format"], "json")
        self.assertEqual(commands["launch"]["output"]["format"], "json")
        self.assertEqual(commands["apply"]["output"]["format"], "json")
        self.assertEqual(commands["resolve"]["output"]["format"], "stdout")
        self.assertIn("MERGE_SERVER_URL", commands["resolve"]["output"]["description"])
        # resolve exit 0 means a resolution was applied, not merely a clean stop
        self.assertIn("resolution applied", commands["resolve"]["exitCodes"]["0"])
        self.assertIn("before a resolution was applied", commands["resolve"]["exitCodes"]["1"])

    def test_shim_python_dependencies_match_manifest_runtime(self):
        # bin/xlmerge.js 固定依赖必须与 manifest runtime.pythonDependencies 逐项一致
        shim = (PACKAGE_ROOT / "bin" / "xlmerge.js").read_text(encoding="utf-8")
        match = re.search(r"const DEPS = \[(.*?)\];", shim, re.S)
        self.assertIsNotNone(match)
        deps = re.findall(r"\"([^\"]+)\"", match.group(1))
        self.assertEqual(deps, self.manifest["runtime"]["pythonDependencies"])
        self.assertEqual(len(deps), 2)

    def test_cli_repo_defaults_to_current_directory(self):
        # 回归：--repo 缺省必须是当前目录（曾误为字面逗号 ","）
        args = build_parser().parse_args(["detect"])
        self.assertEqual(args.repo, ".")
        self.assertNotEqual(args.repo, ",")
        self.assertIsNone(args.runtime_dir)


if __name__ == "__main__":
    unittest.main()