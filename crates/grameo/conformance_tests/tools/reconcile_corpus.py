#!/usr/bin/env python3
"""Replay changed raw fixtures against the frozen full-gate binaries.

Preserves the original report and changed TSVs, snapshots every delta input,
records commands/binary hashes, and publishes a report only after every current
language corpus is fully accounted for and its fingerprints match.
"""
import argparse
import csv
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import time

import full_corpus as gate
import audit_fingerprints


def rows(path):
    with path.open() as source:
        return list(csv.DictReader(source, delimiter="\t"))


def write_rows(path, values):
    with path.open("w", newline="") as target:
        writer = csv.DictWriter(target, fieldnames=values[0].keys(), delimiter="\t", lineterminator="\n")
        writer.writeheader()
        writer.writerows(values)


def inventory(repo, language):
    root = repo / "tests" / ("shell" if language == "washm" else language)
    return {str(path.relative_to(root)): path for path in root.rglob("*")
            if path.is_file() and path.suffix[1:] in gate.EXTENSIONS[language]}


def corpus_digest(current):
    digest = hashlib.sha256()
    for label, path in sorted(current.items()):
        for part in (label.encode(), path.read_bytes()):
            digest.update(len(part).to_bytes(8, "little"))
            digest.update(part)
    return digest.hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    crate = Path(__file__).resolve().parents[2]
    repo = crate.parents[1]
    output = args.output.resolve()
    original = json.loads((output / "report.json").read_text())
    assert original["all_languages"] and original["implementation_unchanged"]
    assert all(batch["exit_code"] == 0 for batch in original["batches"].values())
    expected_hash = original["implementation_end_sha256"]
    assert gate.implementation_hash(crate) == expected_hash
    assert all(v["complete"] and v["mismatches"] == v["resource_mismatches"] == v["unreadable"] == 0
               for v in original["languages"].values())
    prior_archives = sorted(str(path) for path in output.glob("delta-*") if path.is_dir())
    archive = output / f"delta-{time.time_ns()}"
    archive.mkdir()
    working = archive / "working"
    working.mkdir()
    for language, values in original["languages"].items():
        path = output / f"{language}.tsv"
        assert hashlib.sha256(path.read_bytes()).hexdigest() == values["report_sha256"], f"{language} audited TSV changed since report publication"
        shutil.copy2(path, working / path.name)
    fast_fingerprint, fingerprint_provenance = audit_fingerprints.build(archive / "audit-helper", gate.input_fingerprint)
    shutil.copy2(output / "report.json", archive / "report-before.json")
    report = json.loads(json.dumps(original))
    provenance = {"started_unix": time.time(), "original_report": str(archive / "report-before.json"), "prior_delta_archives": prior_archives, "fingerprint_helper": fingerprint_provenance, "rounds": []}
    batches = {language: batch for batch, languages in gate.BATCHES.items() for language in languages}
    for round_index in range(8):
        round_record = {"round": round_index, "languages": {}}
        pending = {}
        for language, batch in batches.items():
            current = inventory(repo, language)
            if corpus_digest(current) == original["languages"][language]["raw_corpus_sha256"]:
                continue  # The original Rust rows were already attested to these exact bytes.
            existing = {gate.unescape(r["label"]): r for r in rows(working / f"{language}.tsv")}
            raw_existing = {label: row for label, row in existing.items() if not label.startswith("preprocessed:")}
            deleted = sorted(raw_existing.keys() - current.keys())
            changed = sorted(label for label, path in current.items()
                             if label not in existing or fast_fingerprint(label, path.read_bytes()) != existing[label]["fingerprint"])
            if not deleted and not changed:
                continue
            pending[language] = (current, existing, deleted, changed, batch)
        if not pending:
            break
        provenance["rounds"].append(round_record)
        for language, (current, existing, deleted, changed, batch) in pending.items():
            round_dir = archive / f"round-{round_index}" / language
            round_dir.mkdir(parents=True)
            shutil.copy2(working / f"{language}.tsv", round_dir / "before.tsv")
            binary_info = original["batches"][batch]
            if "binary_snapshot" in binary_info:
                binary = Path(binary_info["binary_snapshot"])
                assert hashlib.sha256(binary.read_bytes()).hexdigest() == binary_info["binary_sha256"]
            else:  # Older reports predate immutable binary snapshots.
                log_text = (output / f"{batch}.log").read_text()
                binary = Path(re.search(r"Running tests/repository_languages\.rs \(([^\n]+)\)", log_text)[1])
            record = {"deleted": deleted, "replays": [], "binary": str(binary), "binary_sha256": hashlib.sha256(binary.read_bytes()).hexdigest()}
            round_record["languages"][language] = record
            for index, label in enumerate(changed):
                replay_dir = round_dir / f"input-{index}"
                replay_dir.mkdir()
                data = current[label].read_bytes()
                data.decode("utf-8")  # Invalid UTF-8 is a blocking error, never a dropped input.
                snapshot = replay_dir / "input.source"
                snapshot.write_bytes(data)
                transformed = language in ("python", "fortran")
                # Preserve the grammar runner's canonical label for transformed
                # inputs, whose bytes are produced by the Rust frontend helper.
                # Its raw row and a post-run comparison attest the snapshot.
                replay_input = crate / "conformance_tests" / "../../../tests" / language / label if transformed else snapshot
                env = os.environ.copy()
                env.update(RUST_MIN_STACK="67108864", VYBE_PARSER_FULL_CORPUS="1", VYBE_PARSER_AUDIT_DIR=str(replay_dir), VYBE_PARSER_REPLAY_INPUT=str(replay_input))
                command = [str(binary), f"{language}::repository_source_acceptance_capture_and_native_diagnostic_parity", "--exact", "--nocapture", "--test-threads=1"]
                with (replay_dir / "run.log").open("w") as log:
                    result = subprocess.run(command, env=env, cwd="/private/tmp", stdout=log, stderr=subprocess.STDOUT)
                replay_rows = rows(replay_dir / f"{language}.tsv")
                summary = dict(line.split("=", 1) for line in (replay_dir / f"{language}.summary").read_text().splitlines())
                count = 2 if transformed else 1
                assert result.returncode == 0 and len(replay_rows) == count and summary["inputs"] == str(count)
                assert summary["mismatches"] == summary["resource_mismatches"] == summary["unreadable"] == "0"
                row = replay_rows[0]
                assert row["status"] == "parity" and int(row["bytes"]) == len(data)
                assert row["fingerprint"] == gate.input_fingerprint(gate.unescape(row["label"]), data)
                row["label"] = label.replace("\\", "\\\\").replace("\t", "\\t").replace("\n", "\\n").replace("\r", "\\r")
                row["fingerprint"] = gate.input_fingerprint(label, data)
                if transformed:
                    assert current[label].read_bytes() == data, "Transformed source changed during replay"
                    processed = replay_rows[1]
                    assert gate.unescape(processed["label"]) == f"preprocessed:{label}" and processed["status"] == "parity"
                    existing[f"preprocessed:{label}"] = processed
                record["replays"].append({"label": label, "kind": "modified" if label in existing else "added", "source_sha256": hashlib.sha256(data).hexdigest(), "command": command, "exit_code": result.returncode, "canonical_rows": replay_rows})
                (archive / "provenance.json").write_text(json.dumps(provenance, indent=2) + "\n")
                existing[label] = row
            for label in deleted:
                del existing[label]
                existing.pop(f"preprocessed:{label}", None)
            write_rows(working / f"{language}.tsv", [existing[label] for label in sorted(existing)])
    else:
        raise RuntimeError("Corpus did not reach a fixed point after eight delta rounds")
    # A final full scan is mandatory, including languages with no delta.
    for language, values in report["languages"].items():
        audited = rows(working / f"{language}.tsv")
        current = inventory(repo, language)
        digest = corpus_digest(current)
        stable = digest == original["languages"][language]["raw_corpus_sha256"]
        if not stable:
            raw_rows = {gate.unescape(r["label"]): r for r in audited if not r["label"].startswith("preprocessed:")}
            stable = current.keys() == raw_rows.keys() and all(fast_fingerprint(label, path.read_bytes()) == raw_rows[label]["fingerprint"] for label, path in current.items())
        assert stable, f"{language} corpus changed during final verification"
        raw = sum(not r["label"].startswith("preprocessed:") for r in audited)
        values.update(files=raw, inputs=len(audited), accepted=sum(r["reference"] == "accepted" for r in audited), rejected=sum(r["reference"] == "rejected" for r in audited), raw_input_rows=raw, preprocessed_input_rows=len(audited)-raw, complete=True, corpus_unchanged=True, raw_corpus_sha256=digest, report_sha256=hashlib.sha256((working / f"{language}.tsv").read_bytes()).hexdigest(), diagnostic_offset_differences=sum(int(r["diagnostic_offset_difference"]) for r in audited), input_bytes=sum(int(r["bytes"]) for r in audited))
        assert values["accepted"] + values["rejected"] == values["inputs"]
        (working / f"{language}.summary").write_text("".join(f"{key}={values[key]}\n" for key in ("files", "inputs", "accepted", "rejected", "mismatches", "resource_mismatches", "unreadable")))
    # Fresh C-fast SHA scans also catch changes after an earlier language scan.
    assert all(corpus_digest(inventory(repo, language)) == values["raw_corpus_sha256"]
               for language, values in report["languages"].items()), "Corpus changed during final SHA verification"
    assert gate.implementation_hash(crate) == expected_hash
    provenance["finished_unix"] = time.time()
    report.setdefault("incremental_reconciliations", []).append(provenance)
    report["finished_unix"] = provenance["finished_unix"]
    report["gate_passed"] = True
    (archive / "provenance.json").write_text(json.dumps(provenance, indent=2) + "\n")
    # A failed reconciliation never mutates the published audit rows. Mark the
    # old report unready while replacing them, then publish the final report.
    publishing = json.loads(json.dumps(original))
    publishing["gate_passed"] = False
    publishing["publication_in_progress"] = str(archive)
    (output / "report.json").write_text(json.dumps(publishing, indent=2) + "\n")
    for path in working.iterdir():
        shutil.copy2(path, output / path.name)
    (output / "report.json").write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps({"gate_passed": True, "archive": str(archive), "raw_files": sum(v["files"] for v in report["languages"].values()), "inputs": sum(v["inputs"] for v in report["languages"].values())}, indent=2))


if __name__ == "__main__":
    main()
