#!/usr/bin/env python3
"""Compare OLD pest-attested native and NEW native full trees byte-for-byte.

The historical pest gate remains untouched. Added/modified current files first
run through immutable reference binaries. No new pest derives are performed.
"""
import argparse
import csv
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import struct
import subprocess
import time

import audit_fingerprints
import full_corpus as gate
import reconcile_corpus as reconciliation


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def put_string(target, value):
    target.write(struct.pack("<Q", len(value)))
    target.write(value)


def bundle(path, language, sources, expected):
    with path.open("wb") as target:
        target.write(struct.pack("<Q", len(sources)))
        for label, data in sources:
            put_string(target, label.encode())
            put_string(target, data)
            labels = [label, f"preprocessed:{label}"] if language in ("python", "fortran") else [label]
            for current in labels:
                row = expected[current]
                target.write(struct.pack("<BQQ", int(row["reference"] == "accepted"), int(row["bytes"]), int(row["fingerprint"], 16)))


def source_digest(sources):
    checksum = hashlib.sha256()
    for label, data in sources:
        for part in (label.encode(), data):
            checksum.update(struct.pack("<Q", len(part)))
            checksum.update(part)
    return checksum.hexdigest()


def bundled_sources(path, language):
    with path.open("rb") as source:
        count = struct.unpack("<Q", read_exact(source, 8))[0]
        for _ in range(count):
            label = read_exact(source, struct.unpack("<Q", read_exact(source, 8))[0]).decode()
            data = read_exact(source, struct.unpack("<Q", read_exact(source, 8))[0])
            read_exact(source, 34 if language in ("python", "fortran") else 17)
            yield label, data
        assert not source.read(1)


def read_exact(source, count):
    data = bytearray()
    while len(data) < count:
        more = source.read(count - len(data))
        if not more:
            raise EOFError(f"driver exited after {len(data)}/{count} bytes")
        data.extend(more)
    return bytes(data)


def record(source):
    prefix = source.read(8)
    if not prefix:
        return None
    if len(prefix) != 8:
        raise EOFError("partial label prefix")
    label = read_exact(source, struct.unpack("<Q", prefix)[0]).decode()
    size = struct.unpack("<Q", read_exact(source, 8))[0]
    return label, read_exact(source, size)


def decode(data):
    index = 0
    def integer():
        nonlocal index
        value = struct.unpack_from("<Q", data, index)[0]
        index += 8
        return value
    def string():
        nonlocal index
        size = integer()
        value = data[index:index+size].decode()
        index += size
        return value
    source_bytes, fingerprint = integer(), integer()
    accepted = bool(data[index]); index += 1
    outcomes = []
    for backend in range(4):
        size = integer(); end = index + size
        ok = bool(data[index]); index += 1
        if ok:
            consumed = integer()
            outcome = {"accepted": True, "consumed": consumed}
            if backend < 2:
                outcome["captures"] = [(string(), integer(), integer(), integer()) for _ in range(integer())]
            else:
                outcome["logical_steps"] = integer()
        else:
            outcome = {"accepted": False, "kind": string(), "diagnostic": string()}
        assert index == end
        outcomes.append(outcome)
    assert index == len(data)
    return {"bytes": source_bytes, "fingerprint": f"{fingerprint:016x}", "reference_accepted": accepted, "outcomes": outcomes}


def compile_driver(target, deps, parser, generated, helpers, log):
    source = Path(__file__).with_name("native_replay.rs")
    identity = {"parser_sha256": digest(parser), "generated_sha256": digest(generated), "source_sha256": digest(source), "helper_sha256": {name: digest(path) for name, path in helpers.items()}}
    cached = target.with_suffix(".build.json")
    if target.exists() and cached.exists():
        previous = json.loads(cached.read_text())
        if all(previous.get(key) == value for key, value in identity.items()) and digest(target) == previous["driver_sha256"]:
            return previous
    env = os.environ.copy()
    env.update(DEVELOPER_DIR="/Library/Developer/CommandLineTools", VYBE_AUDIT_PYTHON_HELPER=str(helpers["python"]), VYBE_AUDIT_FORTRAN_HELPER=str(helpers["fortran"]))
    command = ["rustc", "--edition=2024", "-O", str(source), "--extern", f"vybe_parser={parser}", "--extern", f"vybe_parser_generated_tests={generated}", "-L", f"dependency={deps}", "-o", str(target)]
    with log.open("w") as output:
        subprocess.run(command, env=env, stdout=output, stderr=subprocess.STDOUT, check=True)
    metadata = {"command": command, "driver_sha256": digest(target), **identity}
    cached.write_text(json.dumps(metadata, indent=2) + "\n")
    return metadata


def compare(language, bundle_path, old, new, grammars, output, steps=10_000_000, depth=4096, source_map=None):
    resource = steps != 10_000_000 or depth != 4096
    stem = f"{language}-steps{steps}-depth{depth}"
    env = os.environ.copy(); env["RUST_MIN_STACK"] = "67108864"
    logs = [(output / f"{stem}-{version}.log").open("wb") for version in ("old", "new")]
    processes = [subprocess.Popen([str(executable), language, str(bundle_path), str(grammars), str(steps), str(depth)], stdout=subprocess.PIPE, stderr=log, env=env) for executable, log in zip((old, new), logs)]
    records, failures, step_failures = [], 0, 0
    try:
        while True:
            before, after = [record(process.stdout) for process in processes]
            if before is None or after is None:
                assert before is after, f"{language}: unequal record counts"
                break
            assert before[0] == after[0]
            label, old_record = before
            new_record = after[1]
            old_value = decode(old_record)
            new_value = decode(new_record)
            if not resource:
                for outcome in old_value["outcomes"]:
                    assert outcome["accepted"] == old_value["reference_accepted"], "OLD native contradicted pest attestation"
                    if outcome["accepted"]:
                        assert outcome["consumed"] == old_value["bytes"]
                    else:
                        assert outcome["kind"] == "Syntax"
            parity = old_record == new_record
            step_delta = any(before.get("logical_steps") != after.get("logical_steps") for before, after in zip(old_value["outcomes"][2:], new_value["outcomes"][2:]) if before["accepted"] and after["accepted"])
            step_failures += step_delta
            if not parity:
                failures += 1
                failure = output / f"{stem}-mismatch-{failures}"
                failure.with_suffix(".old.json").write_text(json.dumps(old_value, indent=2))
                failure.with_suffix(".new.json").write_text(json.dumps(new_value, indent=2))
                failure.with_suffix(".label").write_text(label)
                if source_map is not None:
                    raw_label = label.removeprefix("preprocessed:")
                    failure.with_suffix(".source").write_bytes(source_map[raw_label])
            records.append({"label": label, "bytes": old_value["bytes"], "fingerprint": old_value["fingerprint"], "reference": "accepted" if old_value["reference_accepted"] else "rejected", "old_native_sha256": hashlib.sha256(old_record).hexdigest(), "new_native_sha256": hashlib.sha256(new_record).hexdigest(), "status": "parity" if parity else "mismatch"})
        assert all(process.wait() == 0 for process in processes), f"{language}: driver failure; see logs"
    finally:
        for process in processes:
            if process.poll() is None:
                process.terminate(); process.wait()
        for log in logs: log.close()
    return records, failures, step_failures


def compare_recorded(language, bundle_path, old, new, grammars, output, expected, sources):
    """Compare complete NEW records with cryptographically pinned OLD records."""
    records, failures, step_failures = [], 0, 0
    seen = set()
    with (output / f"{language}-recorded-new.log").open("wb") as log:
        process = subprocess.Popen([str(new), language, str(bundle_path), str(grammars), "10000000", "4096"], stdout=subprocess.PIPE, stderr=log)
        try:
            while True:
                item = record(process.stdout)
                if item is None:
                    break
                label, data = item
                assert label in expected and label not in seen
                seen.add(label)
                oracle = expected[label]
                assert oracle["status"] == "parity" and oracle["old_native_sha256"] == oracle["new_native_sha256"]
                checksum = hashlib.sha256(data).hexdigest()
                parity = checksum == oracle["old_native_sha256"]
                # Digest identity covers the complete captures/errors/step
                # payload, so use its already-attested header on a match.
                value = {"bytes": oracle["bytes"], "fingerprint": oracle["fingerprint"], "reference_accepted": oracle["reference"] == "accepted"} if parity else decode(data)
                if not parity:
                    failures += 1
                    # Recover the old full record only for a concrete mismatch,
                    # retaining reproducible JSON and separate step assertions.
                    failure = output / "recorded-mismatches" / language / str(failures)
                    failure.mkdir(parents=True)
                    raw_label = label.removeprefix("preprocessed:")
                    probe = failure / "input.bundle"
                    bundle(probe, language, [(raw_label, sources[raw_label])], expected)
                    _, _, delta = compare(language, probe, old, new, grammars, failure, source_map=sources)
                    step_failures += delta
                records.append({"label": label, "bytes": value["bytes"], "fingerprint": value["fingerprint"], "reference": "accepted" if value["reference_accepted"] else "rejected", "old_native_sha256": oracle["old_native_sha256"], "new_native_sha256": checksum, "status": "parity" if parity else "mismatch"})
            assert process.wait() == 0
        finally:
            if process.poll() is None:
                process.terminate(); process.wait()
    assert {row["label"] for row in records} == set(expected)
    return records, failures, step_failures


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline", required=True, type=Path)
    parser.add_argument("--candidate-target", required=True, type=Path)
    parser.add_argument("--candidate-crate", type=Path, help="isolated new-engine crate using all baseline grammar texts")
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--languages", nargs="+", choices=gate.EXTENSIONS, default=list(gate.EXTENSIONS))
    parser.add_argument("--limit", type=int, help="smoke only; prevents a full-gate claim")
    parser.add_argument("--resume", action="store_true", help="preserve prior evidence and replay only changed corpora with the same candidate")
    parser.add_argument("--frozen-input-report", type=Path, help="replay immutable bundles from a completed pest-attested native differential instead of moving live inputs")
    parser.add_argument("--recorded-oracle", action="store_true", help="compare NEW complete record hashes with the pinned full-record OLD evidence; rerun OLD only for mismatches and guards")
    args = parser.parse_args()
    baseline, output, candidate_target = args.baseline.resolve(), args.output.resolve(), args.candidate_target.resolve()
    output.mkdir(parents=True, exist_ok=True)
    crate = Path(__file__).resolve().parents[2]; repo = crate.parents[1]
    candidate_crate = args.candidate_crate.resolve() if args.candidate_crate else crate
    source_hash = gate.implementation_hash(candidate_crate)
    metadata = json.loads((baseline / "manifest.json").read_text())
    previous = json.loads((baseline / "attested-inputs/report.json").read_text())
    assert previous["gate_passed"]
    for relative, value in metadata["files"].items():
        assert digest(baseline / relative) == value["sha256"], f"baseline artifact changed: {relative}"
    grammar_drift = {}
    for language, values in previous["languages"].items():
        assert digest(candidate_crate.parents[1] / f"languages/{language}/src/grammar.pest") == values["grammar_sha256"], "candidate grammar changed; requires new pest gate"
        actual = digest(repo / f"languages/{language}/src/grammar.pest")
        if actual != values["grammar_sha256"]:
            grammar_drift[language] = {"baseline_sha256": values["grammar_sha256"], "production_sha256": actual}
    for language in ("python", "fortran"):
        relative = f"sources/languages/{language}/src/walker.rs"
        assert digest(repo / f"languages/{language}/src/walker.rs") == metadata["files"][relative]["sha256"]
    helpers = {}
    for language in ("python", "fortran"):
        candidates = list((baseline / "generated-sources").rglob(f"{language}_preprocessor.rs"))
        assert len({digest(path) for path in candidates}) == 1
        helpers[language] = candidates[0]
    artifacts = [json.loads(line) for line in (candidate_target / "build.jsonl").read_text().splitlines() if line.startswith("{")]
    def artifact(name):
        candidates = [item for item in artifacts if item.get("reason") == "compiler-artifact" and item["target"]["name"] == name and item["profile"]["opt_level"] == "3" and any(path.endswith(".rlib") for path in item["filenames"])]
        item = candidates[-1]
        return Path(next(path for path in item["filenames"] if path.endswith(".rlib")))
    new_parser, new_generated = artifact("vybe_parser"), artifact("vybe_parser_generated_tests")
    old_parser, old_generated = [baseline / metadata[key] for key in ("preferred_parser_artifact", "preferred_generated_artifact")]
    report = {"schema": 1, "scope": "Exact native differential on identical frozen grammars; OLD full captures/consumed/recognition attested against pest; production AST/runtime and drifted production grammars not covered", "started_unix": time.time(), "candidate_crate": str(candidate_crate), "candidate_start_sha256": source_hash, "baseline_implementation_sha256": previous["implementation_end_sha256"], "production_grammar_drift": grammar_drift, "languages": {}, "drivers": {}, "pest_delta_replays": []}
    frozen = None
    if args.frozen_input_report:
        assert not args.resume and not args.limit
        frozen_path = args.frozen_input_report.resolve()
        frozen = json.loads(frozen_path.read_text())
        assert frozen["immutable_bundle_gate_passed"]
        assert frozen["baseline_implementation_sha256"] == previous["implementation_end_sha256"]
        shutil.copy2(frozen_path, output / "frozen-input-report.json")
        report["frozen_input_provenance"] = {"report": str(frozen_path), "report_sha256": digest(frozen_path)}
    if args.recorded_oracle:
        assert frozen is not None
        assert frozen["drivers"]["old"]["source_sha256"] == digest(Path(__file__).with_name("native_replay.rs"))
        report["comparison_mode"] = "NEW full canonical records against SHA256-pinned exact OLD records; OLD replay on mismatches and resource probes"
    archive = None
    if args.resume:
        report = json.loads((output / "report.json").read_text())
        assert report["candidate_end_sha256"] == source_hash, "candidate changed; use a new output directory"
        archive = output / f"reconciliation-{time.time_ns()}"
        archive.mkdir()
        shutil.copy2(output / "report.json", archive / "report-before.json")
        report.setdefault("reconciliations", []).append({"started_unix": time.time(), "archive": str(archive), "languages": []})
        report["gate_passed"] = False
        for language, values in report["languages"].items():
            assert digest(output / f"{language}.bundle") == values["bundle_sha256"]
            assert digest(output / f"{language}.tsv") == values["report_sha256"]
            values["raw_corpus_sha256"] = source_digest(bundled_sources(output / f"{language}.bundle", language))
            values["corpus_unchanged"] = reconciliation.corpus_digest(reconciliation.inventory(repo, language)) == values["raw_corpus_sha256"]
    old, new = output / "old-native", output / "new-native"
    report["drivers"]["old"] = compile_driver(old, baseline / "dependencies", old_parser, old_generated, helpers, output / "old-build.log")
    report["drivers"]["new"] = compile_driver(new, new_parser.parent, new_parser, new_generated, helpers, output / "new-build.log")
    fingerprint, report["fingerprint_helper"] = audit_fingerprints.build(output / "audit-helper", gate.input_fingerprint)
    grammars = baseline / "sources/languages"
    for language in args.languages:
        if args.resume and language in report["languages"] and report["languages"][language]["corpus_unchanged"]:
            print(f"Retaining unchanged {language}", flush=True)
            continue
        if archive:
            report["reconciliations"][-1]["languages"].append(language)
            for suffix in ("bundle", "tsv", "resource.tsv", "resource.bundle"):
                path = output / f"{language}.{suffix}" if not suffix.startswith("resource") else output / f"{language}-{suffix}"
                if path.exists(): shutil.copy2(path, archive / path.name)
        (output / f"{language}-resource.tsv").unlink(missing_ok=True)
        print(f"Preparing {language}", flush=True)
        expected = {gate.unescape(row["label"]): row for row in reconciliation.rows(baseline / f"attested-inputs/{language}.tsv")}
        current = {} if frozen else reconciliation.inventory(repo, language)
        sources = []
        if frozen:
            frozen_root = args.frozen_input_report.resolve().parent
            values = frozen["languages"][language]
            assert digest(frozen_root / f"{language}.bundle") == values["bundle_sha256"]
            assert digest(frozen_root / f"{language}.tsv") == values["report_sha256"]
            sources = list(bundled_sources(frozen_root / f"{language}.bundle", language))
            assert source_digest(sources) == values["raw_corpus_sha256"]
            expected = {row["label"]: row for row in reconciliation.rows(frozen_root / f"{language}.tsv")}
        for label, path in sorted(current.items()):
            data = path.read_bytes(); data.decode()
            if label not in expected or fingerprint(label, data) != expected[label]["fingerprint"]:
                replay = output / "pest-deltas" / language / str(len(report["pest_delta_replays"]))
                replay.mkdir(parents=True)
                snapshot = replay / "input.source"; snapshot.write_bytes(data)
                transformed = language in ("python", "fortran")
                replay_input = crate / "conformance_tests" / "../../../tests" / language / label if transformed else snapshot
                batch = next(batch for batch, langs in gate.BATCHES.items() if language in langs)
                binary = baseline / "reference-binaries" / batch
                env = os.environ.copy(); env.update(RUST_MIN_STACK="67108864", VYBE_PARSER_FULL_CORPUS="1", VYBE_PARSER_REPLAY_INPUT=str(replay_input), VYBE_PARSER_AUDIT_DIR=str(replay))
                command = [str(binary), f"{language}::repository_source_acceptance_capture_and_native_diagnostic_parity", "--exact", "--nocapture"]
                with (replay / "run.log").open("w") as log:
                    subprocess.run(command, env=env, stdout=log, stderr=subprocess.STDOUT, check=True)
                rows = reconciliation.rows(replay / f"{language}.tsv")
                assert len(rows) == (2 if transformed else 1) and all(row["status"] == "parity" for row in rows)
                assert fingerprint(gate.unescape(rows[0]["label"]), data) == rows[0]["fingerprint"]
                rows[0].update(label=label, fingerprint=fingerprint(label, data))
                expected[label] = rows[0]
                if transformed:
                    assert path.read_bytes() == data
                    assert gate.unescape(rows[1]["label"]) == f"preprocessed:{label}"
                    expected[f"preprocessed:{label}"] = rows[1]
                report["pest_delta_replays"].append({"language": language, "label": label, "command": command, "source_sha256": hashlib.sha256(data).hexdigest(), "rows": rows})
            sources.append((label, data))
        if args.limit:
            sources = sources[:args.limit]
        path = output / f"{language}.bundle"
        if frozen:
            shutil.copy2(frozen_root / f"{language}.bundle", path)
        else:
            bundle(path, language, sources, expected)
        print(f"Comparing {language}: {len(sources)} raw inputs", flush=True)
        source_map = dict(sources)
        if args.recorded_oracle:
            records, failures, step_failures = compare_recorded(language, path, old, new, grammars, output, expected, source_map)
        else:
            records, failures, step_failures = compare(language, path, old, new, grammars, output, source_map=source_map)
        expected_count = len(sources) * (2 if language in ("python", "fortran") else 1)
        assert len(records) == expected_count
        with (output / f"{language}.tsv").open("w", newline="") as target:
            writer = csv.DictWriter(target, fieldnames=records[0].keys(), delimiter="\t", lineterminator="\n"); writer.writeheader(); writer.writerows(records)
        resource_sources = [item for item in sources if expected[item[0]]["reference"] == "accepted"][:8]
        resource_cases, resource_failures, resource_step_failures = 0, 0, 0
        if resource_sources:
            probe = output / f"{language}-resource.bundle"; bundle(probe, language, resource_sources, expected)
            for steps, depth in [(value, 4096) for value in (0, 1, 2, 7, 31, 127)] + [(10_000_000, value) for value in (0, 1, 2, 7, 31, 127)]:
                result, count, step_count = compare(language, probe, old, new, grammars, output, steps, depth, source_map)
                resource_cases += len(result); resource_failures += count
                resource_step_failures += step_count
                with (output / f"{language}-resource.tsv").open("a") as target:
                    for row in result: target.write(f"{row['label']}\t{steps}\t{depth}\t{row['old_native_sha256']}\t{row['new_native_sha256']}\t{row['status']}\n")
        sha = source_digest(sources)
        report["languages"][language] = {"raw_files": len(sources), "inputs": len(records), "accepted": sum(row["reference"] == "accepted" for row in records), "rejected": sum(row["reference"] == "rejected" for row in records), "mismatches": failures, "logical_step_mismatches": step_failures, "resource_cases": resource_cases, "resource_mismatches": resource_failures, "resource_logical_step_mismatches": resource_step_failures, "raw_corpus_sha256": sha, "bundle_sha256": digest(path), "report_sha256": digest(output / f"{language}.tsv"), "corpus_unchanged": reconciliation.corpus_digest(reconciliation.inventory(repo, language)) == sha}
        (output / "report.json").write_text(json.dumps(report, indent=2) + "\n")
        print(f"Completed {language}: mismatches={failures} resource_mismatches={resource_failures}", flush=True)
    for language, values in report["languages"].items():
        values["corpus_unchanged"] &= reconciliation.corpus_digest(reconciliation.inventory(repo, language)) == values["raw_corpus_sha256"]
    report["candidate_end_sha256"] = gate.implementation_hash(candidate_crate)
    report["candidate_unchanged"] = report["candidate_end_sha256"] == source_hash
    report["all_languages"] = len(report["languages"]) == 18 and not args.limit
    report["immutable_bundle_gate_passed"] = report["all_languages"] and report["candidate_unchanged"] and all(values["mismatches"] == values["resource_mismatches"] == values["logical_step_mismatches"] == values["resource_logical_step_mismatches"] == 0 for values in report["languages"].values())
    report["gate_passed"] = report["immutable_bundle_gate_passed"] and all(values["corpus_unchanged"] for values in report["languages"].values())
    report["finished_unix"] = time.time()
    (output / "report.json").write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps({"gate_passed": report["gate_passed"], "inputs": sum(v["inputs"] for v in report["languages"].values()), "mismatches": sum(v["mismatches"] for v in report["languages"].values()), "resource_mismatches": sum(v["resource_mismatches"] for v in report["languages"].values())}, indent=2))
    return 0 if report["gate_passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
