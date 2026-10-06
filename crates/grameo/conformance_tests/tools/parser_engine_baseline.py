#!/usr/bin/env python3
"""Build old/current/pest in one release executable on frozen grammar inputs."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    repo = Path(__file__).resolve().parents[4]
    args = argparse.ArgumentParser(description=__doc__)
    args.add_argument("--baseline", type=Path, default=repo / "results/testrunner/parser-full-corpus-refresh/native-baseline")
    args.add_argument("--output", type=Path, default=repo / "results/testrunner/parser-engine-repro")
    args.add_argument("--before-core", type=Path)
    args.add_argument("--candidate-core", type=Path, default=repo / "crates/grameo")
    args.add_argument("--samples", type=int, default=501)
    opts = args.parse_args()
    if opts.samples < 1:
        args.error("samples must be positive")
    opts.output.mkdir(parents=True, exist_ok=True)
    frozen = opts.baseline / "sources/crates/vybe_parser"
    manifest = json.loads((opts.baseline / "manifest.json").read_text())
    for name, record in manifest["files"].items():
        if name.startswith("sources/crates/vybe_parser/src/") or name.startswith("sources/crates/vybe_parser/codegen/src/"):
            assert digest(opts.baseline / name) == record["sha256"], name
    work = Path(tempfile.mkdtemp(prefix="vybe-parser-engine-", dir="/private/tmp"))
    current = opts.candidate_core
    old, new, bench = [work / name for name in ("old", "new", "bench")]
    for source, target in [(opts.before_core or frozen, old), (current, new)]:
        shutil.copytree(source / "src", target / "src")
        shutil.copytree(source / "codegen/src", target / "codegen/src")
    shutil.copy2(current / "Cargo.toml", new / "Cargo.toml")
    shutil.copy2(current / "codegen/Cargo.toml", new / "codegen/Cargo.toml")
    (old / "Cargo.toml").write_text('[package]\nname="vybe_parser_before"\nversion="0.1.0"\nedition="2024"\n[workspace]\nmembers=["codegen"]\n[dependencies]\nunicode-ident="=1.0.24"\n')
    (old / "codegen/Cargo.toml").write_text('[package]\nname="vybe_parser_codegen"\nversion="0.1.0"\nedition="2024"\n[dependencies]\nvybe_parser={package="vybe_parser_before",path=".."}\n')
    examples = bench / "examples"
    examples.mkdir(parents=True)
    shutil.copy2(current / "benchmarks/parser_engine_baseline.rs", examples / "baseline.rs")
    (bench / "Cargo.toml").write_text('[package]\nname="parser_engine_benchmark"\nversion="0.1.0"\nedition="2024"\n[workspace]\n[dependencies]\nvybe_parser={path="../new"}\nvybe_parser_before={path="../old"}\npest="=2.8.6"\npest_derive="=2.8.6"\n')
    env = dict(os.environ)
    env.setdefault("DEVELOPER_DIR", "/Library/Developer/CommandLineTools")
    env["CARGO_TARGET_DIR"] = str(work / "target")
    commands = []

    def run(command, stdout=None):
        commands.append(command)
        subprocess.run(command, cwd="/private/tmp", env=env, stdout=stdout, check=True)

    for lang in ["lua", "js"]:
        relative = f"languages/{lang}/src/grammar.pest"
        grammar = opts.baseline / "sources" / relative
        assert digest(grammar) == digest(repo / relative), f"{lang} grammar changed: choose a matching frozen baseline"
        shutil.copy2(grammar, examples / f"{lang}.pest")
        for core, suffix in [(old, "before_generated"), (new, "generated")]:
            target = examples / f"{lang}_{suffix}.rs"
            run(["cargo", "run", "--offline", "-j", "1", "--release", "--manifest-path", str(core / "Cargo.toml"), "-p", "vybe_parser_codegen", "--bin", "grammargen", "--", str(grammar), str(target)])
            if core == old:
                target.write_text(target.read_text().replace("::vybe_parser::", "::vybe_parser_before::"))
    run(["cargo", "build", "--offline", "-j", "1", "--release", "--manifest-path", str(bench / "Cargo.toml"), "--example", "baseline"])
    binary = work / "target/release/examples/baseline"
    for lang in ["lua", "js"]:
        with (opts.output / f"{lang}.log").open("w") as output:
            run([str(binary), *(["--lua"] if lang == "lua" else []), str(opts.samples)], stdout=output)
    proof = {"baseline_manifest_sha256": digest(opts.baseline / "manifest.json"), "samples": opts.samples, "warmups": 10, "workdir": str(work), "binary_sha256": digest(binary), "commands": commands, "files": {str(path.relative_to(work)): digest(path) for path in sorted(work.rglob("*.rs")) if "target" not in path.relative_to(work).parts}}
    (opts.output / "proof.json").write_text(json.dumps(proof, indent=2) + "\n")
    print(f"Benchmarks and provenance: {opts.output}")


if __name__ == "__main__":
    main()
