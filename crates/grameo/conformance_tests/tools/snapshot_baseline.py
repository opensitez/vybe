#!/usr/bin/env python3
"""Archive the pest-attested native dependency closure for later native replay."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import time

import full_corpus as gate


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    output = args.output.resolve()
    report = json.loads((output / "report.json").read_text())
    assert report["gate_passed"] and report["implementation_unchanged"]
    crate = Path(__file__).resolve().parents[2]
    repo = crate.parents[1]
    assert gate.implementation_hash(crate) == report["implementation_end_sha256"]
    destination = output / "native-baseline"
    destination.mkdir()
    manifest = {"schema": 1, "created_unix": time.time(), "scope": "Frozen native libraries and sources attested by the completed pest capture/recognition gate; no new comparison verdicts", "implementation_sha256": report["implementation_end_sha256"], "gate_report_sha256": hashlib.sha256((output / "report.json").read_bytes()).hexdigest(), "rustc": subprocess.run(["rustc", "--version", "--verbose"], capture_output=True, text=True, check=True).stdout, "files": {}, "corpora": {language: {key: values[key] for key in ("files", "inputs", "raw_corpus_sha256", "report_sha256", "grammar_sha256")} for language, values in report["languages"].items()}}

    def archive(source, relative):
        target = destination / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
        original = hashlib.sha256(source.read_bytes()).hexdigest()
        assert hashlib.sha256(target.read_bytes()).hexdigest() == original
        manifest["files"][str(relative)] = {"bytes": target.stat().st_size, "sha256": original, "original": str(source)}

    deps = crate / "conformance_tests/target/release/deps"
    for path in sorted(deps.iterdir()):
        if path.is_file() and path.suffix in (".rlib", ".rmeta", ".dylib", ".so", ".d"):
            archive(path, Path("dependencies") / path.name)
    # Metadata records the actual feature sets and dependency artifact IDs.
    fingerprints = crate / "conformance_tests/target/release/.fingerprint"
    for path in sorted(fingerprints.rglob("*.json")):
        archive(path, Path("cargo-fingerprints") / path.relative_to(fingerprints))
    generated_candidates = []
    for path in deps.glob("libvybe_parser_generated_tests-*.rlib"):
        identifier = path.stem.removeprefix("lib")
        metadata = json.loads((fingerprints / identifier / "lib-vybe_parser_generated_tests.json").read_text())
        if "all-grammars" in metadata["features"]:
            generated_candidates.append((path, metadata))
    generated, metadata = max(generated_candidates, key=lambda item: item[0].stat().st_mtime)
    parser_candidates = []
    for path in deps.glob("libvybe_parser-*.rlib"):
        identifier = path.stem.removeprefix("lib")
        parser_metadata = json.loads((fingerprints / identifier / "lib-vybe_parser.json").read_text())
        if parser_metadata["profile"] == metadata["profile"]:
            parser_candidates.append(path)
    native = max(parser_candidates, key=lambda path: path.stat().st_mtime)
    manifest["preferred_generated_artifact"] = str(Path("dependencies") / generated.name)
    manifest["preferred_parser_artifact"] = str(Path("dependencies") / native.name)
    build = crate / "conformance_tests/target/release/build"
    for pattern in ("vybe_parser_generated_tests-*/out/*.rs", "conformance_tests-*/out/*.rs"):
        for path in sorted(build.glob(pattern)):
            archive(path, Path("generated-sources") / path.relative_to(build))
    for root, directories, files in os.walk(crate):
        directories[:] = sorted(name for name in directories if name not in ("target", "__pycache__", ".git"))
        for name in sorted(files):
            path = Path(root) / name
            relative = path.relative_to(crate)
            if path.suffix in (".rs", ".pest", ".grammar", ".toml", ".lock"):
                archive(path, Path("sources/crates/vybe_parser") / relative)
    for path in sorted((repo / "languages").glob("*/src/grammar.pest")):
        archive(path, Path("sources") / path.relative_to(repo))
    for language in ("python", "fortran"):
        path = repo / f"languages/{language}/src/walker.rs"
        archive(path, Path("sources") / path.relative_to(repo))
    for language, values in report["languages"].items():
        assert hashlib.sha256((output / f"{language}.tsv").read_bytes()).hexdigest() == values["report_sha256"]
        archive(output / f"{language}.tsv", Path("attested-inputs") / f"{language}.tsv")
    archive(output / "report.json", Path("attested-inputs/report.json"))
    for batch, values in report["batches"].items():
        binary = Path(values["binary_snapshot"])
        assert hashlib.sha256(binary.read_bytes()).hexdigest() == values["binary_sha256"]
        archive(binary, Path("reference-binaries") / batch)
    assert gate.implementation_hash(crate) == manifest["implementation_sha256"]
    (destination / "manifest.json").write_text(json.dumps(manifest, indent=2) + "\n")
    print(json.dumps({"snapshot": str(destination), "files": len(manifest["files"]), "bytes": sum(v["bytes"] for v in manifest["files"].values()), "implementation_sha256": manifest["implementation_sha256"]}, indent=2))


if __name__ == "__main__":
    main()
