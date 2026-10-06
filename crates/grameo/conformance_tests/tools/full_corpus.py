#!/usr/bin/env python3
"""Run isolated full-language parser parity batches and save complete accounting.

Usage: python3 crates/grameo/conformance_tests/tools/full_corpus.py \
  --output results/testrunner/parser-full-corpus-final

A successful gate requires all eighteen summaries, no skipped/unreadable input,
no resource or semantic mismatch, and unchanged matcher sources during the run.
This checks grammar captures and recognition, not frontend AST/runtime parity.
"""
import argparse
import csv
import hashlib
import json
import os
from pathlib import Path
import subprocess
import re
import shutil
import time

import audit_fingerprints

EXTENSIONS = {"js": ("js",), "php": ("php",), "python": ("py",), "c": ("c",), "csharp": ("cs",), "java": ("java",), "dart": ("dart",), "go": ("go",), "kotlin": ("kt",), "lua": ("lua",), "powershell": ("ps1",), "ruby": ("rb",), "cobol": ("cob",), "fortran": ("f", "f90"), "pascal": ("pas",), "vb": ("vb",), "washm": ("sh",), "wast": ("wat", "wast")}

BATCHES = {
    "corpus-scripting": ["js", "php", "python"],
    "corpus-c-family": ["c", "csharp", "java"],
    "corpus-jvm-go-dart": ["dart", "go", "kotlin"],
    "corpus-dynamic": ["lua", "powershell", "ruby"],
    "corpus-business": ["cobol", "fortran", "pascal"],
    "corpus-vb-shell-wast": ["vb", "washm", "wast"],
}


def unescape(value):
    result = []
    index = 0
    while index < len(value):
        if value[index] == "\\":
            index += 1
            result.append({"t": "\t", "n": "\n", "r": "\r", "\\": "\\"}[value[index]])
        else:
            result.append(value[index])
        index += 1
    return "".join(result)


def input_fingerprint(label, source):
    value = 0xcbf29ce484222325
    for part in (label.encode(), source):
        for byte in len(part).to_bytes(8, "little") + part:
            value = ((value ^ byte) * 0x100000001b3) & 0xffffffffffffffff
    return f"{value:016x}"


def stable_corpus(repo, language, rows, fingerprint=input_fingerprint):
    root = repo / "tests" / ("shell" if language == "washm" else language)
    files = {str(path.relative_to(root)): path for path in root.rglob("*")
             if path.is_file() and path.suffix[1:] in EXTENSIONS[language]}
    raw = {unescape(row["label"]): row for row in rows
           if not row["label"].startswith("preprocessed:")}
    if files.keys() != raw.keys():
        return False, None
    digest = hashlib.sha256()
    stable = True
    for label, path in sorted(files.items()):
        data = path.read_bytes()
        stable = stable and fingerprint(label, data) == raw[label]["fingerprint"]
        for part in (label.encode(), data):
            digest.update(len(part).to_bytes(8, "little"))
            digest.update(part)
    return stable, digest.hexdigest()


def implementation_hash(crate):
    digest = hashlib.sha256()
    roots = [crate / "src", crate / "codegen/src", crate / "conformance_tests/tests", crate / "generated_tests/src"]
    repo = crate.parents[1]
    paths = [path for root in roots for path in root.rglob("*.rs")]
    paths += list((repo / "languages").glob("*/src/grammar.pest"))
    paths += [repo / f"languages/{language}/src/walker.rs" for language in ("python", "fortran")]
    for path in sorted(paths):
        data = path.read_bytes()
        for item in (str(path.relative_to(repo)).encode(), data):
            digest.update(len(item).to_bytes(8, "little"))
            digest.update(item)
    return digest.hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--batches", nargs="+", choices=BATCHES, default=list(BATCHES))
    args = parser.parse_args()
    crate = Path(__file__).resolve().parents[2]
    repo = crate.parents[1]
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=True)
    fingerprint, fingerprint_provenance = audit_fingerprints.build(output / "audit-helper", input_fingerprint)
    env = os.environ.copy()
    env.pop("VYBE_PARSER_REPLAY_INPUT", None)
    env.update(DEVELOPER_DIR="/Library/Developer/CommandLineTools",
               RUST_MIN_STACK="67108864", VYBE_PARSER_FULL_CORPUS="1",
               VYBE_PARSER_AUDIT_DIR=str(output))
    report = {"schema": 1, "scope": "grammar captures and recognition; not frontend AST/runtime",
              "implementation_start_sha256": implementation_hash(crate),
              "native_limits": {"max_steps": 10000000, "max_rule_depth": 4096},
              "negative_mutations": "run ordinary sampled feature batches separately",
              "started_unix": time.time(), "batches": {}, "languages": {}}
    report["fingerprint_helper"] = fingerprint_provenance
    # Stale summaries must not allow a crashed or failed batch to look complete.
    for batch in args.batches:
        for language in BATCHES[batch]:
            for suffix in ("summary", "tsv", "unreadable.tsv"):
                (output / f"{language}.{suffix}").unlink(missing_ok=True)
            for suffix in ("source", "failure"):
                for artifact in output.glob(f"{language}-*.{suffix}"):
                    artifact.unlink()
        command = ["cargo", "test", "--manifest-path", str(crate / "conformance_tests/Cargo.toml"),
                   "--offline", "--release", "-j", "1", "--test", "repository_languages", "--features", batch,
                   "--", "--nocapture", "--test-threads=1"]
        started = time.monotonic()
        print(f"Running {batch}; log={output / (batch + '.log')}", flush=True)
        with (output / f"{batch}.log").open("w") as log:
            result = subprocess.run(command, cwd="/private/tmp", env=env, stdout=log, stderr=subprocess.STDOUT)
        report["batches"][batch] = {"exit_code": result.returncode, "seconds": time.monotonic() - started,
                                    "command": command}
        matched = re.search(r"Running tests/repository_languages\.rs \(([^\n]+)\)", (output / f"{batch}.log").read_text())
        if matched:
            binary = Path(matched[1])
            frozen = output / "binaries" / batch
            frozen.parent.mkdir(exist_ok=True)
            shutil.copy2(binary, frozen)
            report["batches"][batch].update(binary_snapshot=str(frozen), binary_sha256=hashlib.sha256(frozen.read_bytes()).hexdigest())
        for language in BATCHES[batch]:
            summary = output / f"{language}.summary"
            if not summary.exists():
                report["languages"][language] = {"complete": False}
                continue
            values = dict(line.split("=", 1) for line in summary.read_text().splitlines())
            values = {key: int(value) for key, value in values.items()}
            rows = list(csv.DictReader((output / f"{language}.tsv").open(), delimiter="\t"))
            values["raw_input_rows"] = sum(not row["label"].startswith("preprocessed:") for row in rows)
            values["preprocessed_input_rows"] = len(rows) - values["raw_input_rows"]
            values["complete"] = len(rows) == values["inputs"] and values["raw_input_rows"] + values["unreadable"] == values["files"]
            values["corpus_unchanged"], values["raw_corpus_sha256"] = stable_corpus(repo, language, rows, fingerprint)
            values["report_sha256"] = hashlib.sha256((output / f"{language}.tsv").read_bytes()).hexdigest()
            values["grammar_sha256"] = hashlib.sha256((repo / f"languages/{language}/src/grammar.pest").read_bytes()).hexdigest()
            values["diagnostic_offset_differences"] = sum(int(row["diagnostic_offset_difference"]) for row in rows)
            values["input_bytes"] = sum(int(row["bytes"]) for row in rows)
            report["languages"][language] = values
        (output / "report.json").write_text(json.dumps(report, indent=2) + "\n")
    # Recheck early batches at the end: another session can add/edit fixtures
    # while later languages are running. New inputs must never be silently omitted.
    for language, values in report["languages"].items():
        if values.get("complete"):
            with (output / f"{language}.tsv").open() as source:
                rows = list(csv.DictReader(source, delimiter="\t"))
            stable, digest = stable_corpus(repo, language, rows, fingerprint)
            values["corpus_unchanged"] = values["corpus_unchanged"] and stable and digest == values["raw_corpus_sha256"]
    report["implementation_end_sha256"] = implementation_hash(crate)
    report["implementation_unchanged"] = report["implementation_start_sha256"] == report["implementation_end_sha256"]
    report["finished_unix"] = time.time()
    report["all_languages"] = len(report["languages"]) == 18
    report["gate_passed"] = report["all_languages"] and report["implementation_unchanged"] and all(
        entry["exit_code"] == 0 for entry in report["batches"].values()) and all(
        entry.get("complete") and entry.get("corpus_unchanged") and entry["mismatches"] == 0 and entry["unreadable"] == 0
        for entry in report["languages"].values())
    (output / "report.json").write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps(report, indent=2), flush=True)
    return 0 if report["gate_passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
