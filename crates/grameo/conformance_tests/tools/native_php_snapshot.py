#!/usr/bin/env python3
"""Replay a fixed, pest-attested PHP snapshot with a new native parser."""
import argparse
import csv
import hashlib
import json
from pathlib import Path
import shutil
import time

import full_corpus as gate
import native_differential as native
import reconcile_corpus as reconciliation


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    for option in ("attestation", "old-crate", "old-target", "candidate-crate", "candidate-target", "bundle-report", "reference-baseline", "output"):
        parser.add_argument("--" + option, required=True, type=Path)
    args = parser.parse_args()
    out = args.output.resolve(); out.mkdir(parents=True, exist_ok=True)
    evidence = json.loads(args.attestation.read_text())
    assert evidence["snapshot_gate_passed"]
    assert gate.implementation_hash(args.old_crate) == evidence["implementation_end_sha256"]
    assert native.digest(args.attestation.parent / "php-reference-native") == evidence["binary_sha256"]
    grammar = args.candidate_crate.parents[1] / "languages/php/src/grammar.pest"
    assert native.digest(grammar) == evidence["grammar_snapshot_sha256"]
    proof = json.loads(args.bundle_report.read_text()); values = proof["languages"]["php"]
    assert proof["immutable_bundle_gate_passed"]
    bundle_source = args.bundle_report.parent / "php.bundle"
    assert native.digest(bundle_source) == values["bundle_sha256"]
    sources = list(native.bundled_sources(bundle_source, "php"))
    assert native.source_digest(sources) == values["raw_corpus_sha256"] == evidence["language"]["raw_corpus_sha256"]
    rows = list(csv.DictReader((args.attestation.parent / "php.tsv").open(), delimiter="\t"))
    assert native.digest(args.attestation.parent / "php.tsv") == evidence["language"]["report_sha256"]
    expected = {gate.unescape(row["label"]): row for row in rows}
    assert len(rows) == len(sources)
    native.bundle(out / "php.bundle", "php", sources, expected)
    helpers = {}
    for language in ("python", "fortran"):
        matches = list((args.old_target / "release/build").rglob(f"{language}_preprocessor.rs"))
        if not matches:
            matches = list((args.reference_baseline / "generated-sources").rglob(f"{language}_preprocessor.rs"))
        assert len({native.digest(path) for path in matches}) == 1
        helpers[language] = matches[0].resolve()
    old_deps = args.old_target / "release/deps"
    # Cargo's release runtime profile, distinct from the opt-level-zero build dependency.
    runtime = [path.parent.name.removeprefix("vybe_parser-") for path in (args.old_target / "release/.fingerprint").glob("vybe_parser-*/lib-vybe_parser.json") if json.loads(path.read_text())["profile"] == 2040997289075261528]
    assert len(runtime) == 1
    old_parser = old_deps / f"libvybe_parser-{runtime[0]}.rlib"
    generated = list(old_deps.glob("libvybe_parser_generated_tests-*.rlib")); assert len(generated) == 1
    artifacts = [json.loads(line) for line in (args.candidate_target / "build.jsonl").read_text().splitlines() if line.startswith("{")]
    def artifact(name):
        matches = [item for item in artifacts if item.get("reason") == "compiler-artifact" and item["target"]["name"] == name and item["profile"]["opt_level"] == "3"]
        return Path(next(path for path in matches[-1]["filenames"] if path.endswith(".rlib")))
    new_parser, new_generated = artifact("vybe_parser"), artifact("vybe_parser_generated_tests")
    report = {"schema": 1, "scope": "Exact OLD pest-attested PHP native vs NEW native on one immutable grammar/input snapshot; captures, consumed, recognition, errors, logical steps and resource guards. Not frontend AST/runtime.", "attestation_sha256": native.digest(args.attestation), "grammar_sha256": native.digest(grammar), "candidate_start_sha256": gate.implementation_hash(args.candidate_crate), "raw_corpus_sha256": native.source_digest(sources), "drivers": {}}
    for name, deps, parser_lib, gen_lib in (("old", old_deps, old_parser, generated[0]), ("new", new_parser.parent, new_parser, new_generated)):
        report["drivers"][name] = native.compile_driver(out / f"{name}-native", deps, parser_lib, gen_lib, helpers, out / f"{name}-build.log")
    source_map = dict(sources)
    records, failures, steps = native.compare("php", out / "php.bundle", out / "old-native", out / "new-native", grammar.parents[2], out, source_map=source_map)
    assert len(records) == len(sources)
    with (out / "php.tsv").open("w", newline="") as target:
        writer = csv.DictWriter(target, fieldnames=records[0].keys(), delimiter="\t", lineterminator="\n"); writer.writeheader(); writer.writerows(records)
    probes = [item for item in sources if expected[item[0]]["reference"] == "accepted"][:8]
    native.bundle(out / "php-resource.bundle", "php", probes, expected)
    resource_cases = resource_failures = resource_steps = 0
    for max_steps, depth in [(value, 4096) for value in (0, 1, 2, 7, 31, 127)] + [(10_000_000, value) for value in (0, 1, 2, 7, 31, 127)]:
        results, count, delta = native.compare("php", out / "php-resource.bundle", out / "old-native", out / "new-native", grammar.parents[2], out, max_steps, depth, source_map)
        resource_cases += len(results); resource_failures += count; resource_steps += delta
    report.update(inputs=len(records), accepted=sum(row["reference"] == "accepted" for row in records), rejected=sum(row["reference"] == "rejected" for row in records), mismatches=failures, logical_step_mismatches=steps, resource_cases=resource_cases, resource_mismatches=resource_failures, resource_logical_step_mismatches=resource_steps, candidate_end_sha256=gate.implementation_hash(args.candidate_crate), bundle_sha256=native.digest(out / "php.bundle"), report_sha256=native.digest(out / "php.tsv"), finished_unix=time.time())
    report["candidate_unchanged"] = report["candidate_start_sha256"] == report["candidate_end_sha256"]
    repo = Path(__file__).resolve().parents[4]
    report["live_corpus_matches"] = reconciliation.corpus_digest(reconciliation.inventory(repo, "php")) == report["raw_corpus_sha256"]
    report["live_grammar_sha256"] = native.digest(repo / "languages/php/src/grammar.pest")
    report["immutable_snapshot_gate_passed"] = report["candidate_unchanged"] and failures == steps == resource_failures == resource_steps == 0
    report["live_php_gate_passed"] = report["immutable_snapshot_gate_passed"] and report["live_corpus_matches"] and report["live_grammar_sha256"] == report["grammar_sha256"]
    shutil.copy2(args.attestation, out / "pest-attestation.json")
    shutil.copy2(grammar, out / "grammar.pest")
    (out / "report.json").write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps(report, indent=2), flush=True)
    return 0 if report["immutable_snapshot_gate_passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
