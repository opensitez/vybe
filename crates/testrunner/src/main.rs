//! Vybe test runner.
//!
//! Two jobs:
//!
//!   testrunner extract <rust-test-file>...   native Rust tests → standalone sources
//!   testrunner run <path>...                 run those sources, verdict = exit code
//!
//! It links nothing from Vybe and drives the already-built `vybex` as a
//! subprocess, so a compiler edit never rebuilds it — which is the whole point.
//! It also means every test file is reachable with vybex's own instrumentation
//! (`-g`, `--dap-port`, `--dump-ast`, `-t`) and with the language's real
//! runtime, neither of which `cargo test` can offer.

mod emit;
mod extract;
mod run;
mod rustlit;

use anyhow::{Context, Result};
use rayon::prelude::*;
use std::path::{Path, PathBuf};
use walkdir::WalkDir;

fn main() -> Result<()> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    match args.first().map(String::as_str) {
        Some("extract") => cmd_extract(&args[1..]),
        Some("run") => cmd_run(&args[1..]),
        _ => {
            eprintln!(
                "vybe testrunner\n\n\
                 Usage:\n  \
                 testrunner extract <rust-test-file>... [--out DIR] [--lang NAME]\n  \
                 testrunner run <path>... [--vybex PATH] [--runtime CMD] [-j N] [--verbose]\n"
            );
            std::process::exit(2)
        }
    }
}

fn flag<'a>(args: &'a [String], name: &str) -> Option<&'a str> {
    args.iter()
        .position(|a| a == name)
        .and_then(|i| args.get(i + 1))
        .map(String::as_str)
}

fn positionals(args: &[String]) -> Vec<&String> {
    let mut out = Vec::new();
    let mut skip_next = false;
    for arg in args {
        if skip_next {
            skip_next = false;
            continue;
        }
        if arg.starts_with("--") || arg == "-j" {
            skip_next = !matches!(arg.as_str(), "--verbose");
            continue;
        }
        out.push(arg);
    }
    out
}

// ── extract ─────────────────────────────────────────────────────────────────

fn cmd_extract(args: &[String]) -> Result<()> {
    let out_root = PathBuf::from(flag(args, "--out").unwrap_or("tests"));
    let files = positionals(args);
    anyhow::ensure!(!files.is_empty(), "no input files");

    let mut total = 0usize;
    let mut unpairable: Vec<(String, String)> = Vec::new();

    for input in files {
        let path = Path::new(input);
        let lang = flag(args, "--lang")
            .map(str::to_string)
            .or_else(|| language_of(path))
            .with_context(|| format!("cannot tell the language of {input}"))?;
        let ext = emit::extension(&lang)
            .with_context(|| format!("no source extension known for `{lang}`"))?;

        let text = std::fs::read_to_string(path)
            .with_context(|| format!("reading {input}"))?;
        let cases = extract::cases_in_file(&text)
            .with_context(|| format!("parsing {input}"))?;
        if cases.is_empty() {
            eprintln!("  {input}: no macro cases found");
            continue;
        }

        let category = path
            .file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or("misc")
            .trim_start_matches("test_")
            .to_string();
        let dir = out_root.join(&lang).join(&category);
        std::fs::create_dir_all(&dir)?;
        let harness = emit::harness_body(&lang)?;

        for case in &cases {
            round_trip_check(case, input)?;

            let slug = format!("{lang}/{category}/{}", case.name);
            let emitted = match lang.as_str() {
                "go" => emit::go::emit(case, input, &slug, &harness),
                other => anyhow::bail!("no emitter for `{other}` yet"),
            };
            if let emit::go::Pairing::Unpairable(reason) = &emitted.pairing {
                unpairable.push((slug.clone(), reason.clone()));
            }
            std::fs::write(dir.join(format!("{}.{ext}", case.name)), &emitted.text)?;
            total += 1;
        }
        println!("{input} → {} case(s) in {}", cases.len(), dir.display());
    }

    println!("\nextracted {total} file(s)");
    if unpairable.is_empty() {
        println!("all cases paired 1:1 into assertions");
    } else {
        println!("{} case(s) could NOT be paired:", unpairable.len());
        for (slug, reason) in &unpairable {
            println!("  {slug}: {reason}");
        }
    }
    Ok(())
}

/// Prove the decode was lossless before anything is written. Go and C sources
/// arrive as escaped Rust literals with backtick struct tags nested inside
/// them; a wrong unescape corrupts a program silently, and silently is the
/// failure mode that would poison the whole migration.
fn round_trip_check(case: &extract::Case, origin: &str) -> Result<()> {
    let re_escaped = rustlit::escape(&case.source);
    let (decoded, _) = rustlit::scan(format!("\"{re_escaped}\"").as_bytes(), 0)?;
    anyhow::ensure!(
        decoded == case.source,
        "escape round-trip lost data for `{}` in {origin}",
        case.name
    );
    Ok(())
}

fn language_of(path: &Path) -> Option<String> {
    let mut parts = path.components().map(|c| c.as_os_str().to_string_lossy());
    while let Some(part) = parts.next() {
        if part == "languages" {
            return parts.next().map(|s| s.into_owned());
        }
    }
    None
}

// ── run ─────────────────────────────────────────────────────────────────────

fn cmd_run(args: &[String]) -> Result<()> {
    let vybex = PathBuf::from(flag(args, "--vybex").unwrap_or("target/debug/vybex"));
    let runtime: Option<Vec<String>> = flag(args, "--runtime")
        .map(|cmd| cmd.split_whitespace().map(str::to_string).collect());
    let verbose = args.iter().any(|a| a == "--verbose");
    if let Some(jobs) = flag(args, "-j").and_then(|n| n.parse().ok()) {
        rayon::ThreadPoolBuilder::new().num_threads(jobs).build_global().ok();
    }

    let roots = positionals(args);
    anyhow::ensure!(!roots.is_empty(), "no test paths given");
    let files = collect(&roots);
    anyhow::ensure!(!files.is_empty(), "no test files found");

    let started = std::time::Instant::now();
    let results: Vec<(PathBuf, run::Outcome)> = files
        .par_iter()
        .map(|file| {
            let text = std::fs::read_to_string(file).unwrap_or_default();
            let mode = run::Mode::of(&text);
            let outcome = match &runtime {
                Some(cmd) => run::run_foreign(&cmd[0], &cmd[1..], file, mode),
                None => run::run_case(&vybex, file, mode),
            };
            (file.clone(), outcome)
        })
        .collect();

    let mut failed = Vec::new();
    for (file, outcome) in &results {
        if !outcome.pass {
            failed.push((file, outcome));
        }
    }

    for (file, outcome) in &failed {
        println!("FAIL {}", file.display());
        println!("     {}", run::failure_line(&outcome.output));
        if verbose {
            for line in outcome.output.lines() {
                println!("       | {line}");
            }
        }
    }

    let total = results.len();
    let passed = total - failed.len();
    println!(
        "\n{passed}/{total} passed, {} failed in {:.1}s",
        failed.len(),
        started.elapsed().as_secs_f64()
    );
    if failed.is_empty() { Ok(()) } else { std::process::exit(1) }
}

fn collect(roots: &[&String]) -> Vec<PathBuf> {
    let mut files = Vec::new();
    for root in roots {
        let path = Path::new(root.as_str());
        if path.is_file() {
            files.push(path.to_path_buf());
            continue;
        }
        for entry in WalkDir::new(path).into_iter().flatten() {
            let p = entry.path();
            if p.is_file() && is_test_source(p) {
                files.push(p.to_path_buf());
            }
        }
    }
    files.sort();
    files
}

/// A test file is one that carries the extraction header. Anything else in the
/// tree — the harness included — is not a test.
fn is_test_source(path: &Path) -> bool {
    let Ok(text) = std::fs::read_to_string(path) else {
        return false;
    };
    text.lines().take(5).any(|l| l.contains("vybe-test:"))
}
