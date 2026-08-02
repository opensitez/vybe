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
mod model;
mod pool;
mod report;
mod run;
mod rustlit;

use anyhow::{Context, Result};
use indicatif::{ProgressBar, ProgressStyle};
use rayon::prelude::*;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicUsize, Ordering};
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
                 testrunner run <path>... [options]\n\n\
                 run options:\n  \
                 --vybex PATH    vybex binary to drive      (default target/debug/vybex)\n  \
                 --runtime CMD   run under something else   (\"go run\", \"node\", \"python3\")\n  \
                 --cold          one fresh process per test (default: warm workers)\n  \
                 -j N            parallel workers           (default: CPU count)\n  \
                 --timeout SECS  per-test deadline          (default 30)\n  \
                 --results DIR   report directory           (default results/testrunner)\n  \
                 --verbose       stream every failure as it happens (default: report only)\n\n\
                 A path may be a directory or a single test file. `run` writes a\n\
                 timestamped JSON report and diffs it against the previous run of\n\
                 the same runtime, naming regressions and newly-passing tests.\n"
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
            skip_next = !matches!(arg.as_str(), "--verbose" | "--cold");
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
    // Warm workers are the default; `--cold` forces one fresh process per test,
    // which is the only way to prove a reset is not leaking state between them.
    let cold = args.iter().any(|a| a == "--cold");
    let timeout: u64 = flag(args, "--timeout").and_then(|n| n.parse().ok()).unwrap_or(30);
    let results_dir = PathBuf::from(flag(args, "--results").unwrap_or("results/testrunner"));
    if let Some(jobs) = flag(args, "-j").and_then(|n| n.parse().ok()) {
        rayon::ThreadPoolBuilder::new().num_threads(jobs).build_global().ok();
    }

    let roots = positionals(args);
    anyhow::ensure!(!roots.is_empty(), "no test paths given");
    let files = collect(&roots);
    anyhow::ensure!(!files.is_empty(), "no test files found");

    let threads = rayon::current_num_threads();
    let under = match &runtime {
        Some(cmd) => cmd.join(" "),
        None => vybex.display().to_string(),
    };
    eprintln!(
        "[testrunner] {} test(s) · {threads} {} workers · {timeout}s timeout · under `{under}`",
        files.len(),
        if runtime.is_some() || cold { "cold" } else { "warm" },
    );

    // One process per test, `threads` of them at a time. Without a live bar a
    // long run is indistinguishable from a hung one.
    let bar = ProgressBar::new(files.len() as u64);
    bar.set_style(
        ProgressStyle::with_template(
            "{bar:30} {pos}/{len} {msg} [{per_sec}, {elapsed_precise} elapsed, eta {eta}]",
        )
        .unwrap_or_else(|_| ProgressStyle::default_bar()),
    );
    let passed = AtomicUsize::new(0);
    let failed = AtomicUsize::new(0);

    let started = std::time::Instant::now();
    let note = |exec: &model::TestExecution| {
        if exec.result == model::TestResult::Pass {
            passed.fetch_add(1, Ordering::Relaxed);
        } else {
            failed.fetch_add(1, Ordering::Relaxed);
            // Quiet by default. A failure is a data point, not a message: the
            // live ✗ count says how many, the console report names the first
            // few, and the JSON holds all of them. Streaming every one turns a
            // 759-failure run into 1,518 lines of scrollback.
            if verbose {
                bar.suspend(|| {
                    println!("FAIL {}", exec.slug());
                    println!("     {}", exec.message);
                    println!("     {}", exec.path.display());
                });
            }
        }
        let ok = passed.load(Ordering::Relaxed);
        let bad = failed.load(Ordering::Relaxed);
        bar.set_message(format!("✓{ok} ✗{bad}"));
        bar.inc(1);
        // indicatif draws nothing when stderr is not a terminal (piped to a
        // file, captured by CI). Keep reporting regardless.
        if bar.is_hidden() && (ok + bad) % 250 == 0 {
            eprintln!(
                "[testrunner] {}/{} · ✓{ok} ✗{bad} · {:.0}/s",
                ok + bad,
                files.len(),
                (ok + bad) as f64 / started.elapsed().as_secs_f64().max(0.001),
            );
        }
    };

    let executions: Vec<model::TestExecution> = match &runtime {
        // A foreign runtime has no warm mode — one process per test is all
        // `go run` / `python3` / `node` offer.
        Some(cmd) => files
            .par_iter()
            .map(|file| {
                let text = std::fs::read_to_string(file).unwrap_or_default();
                let mode = run::Mode::of(&text);
                let outcome = run::run_foreign(&cmd[0], &cmd[1..], file, mode, timeout);
                let (language, category, name) = model::identify(file);
                let exec = model::TestExecution {
                    path: file.clone(),
                    language,
                    category,
                    name,
                    result: outcome.result,
                    message: outcome.message,
                    duration_ms: outcome.duration_ms,
                };
                note(&exec);
                exec
            })
            .collect(),
        None if cold => files
            .par_iter()
            .map(|file| {
                let text = std::fs::read_to_string(file).unwrap_or_default();
                let mode = run::Mode::of(&text);
                let outcome = run::run_case(&vybex, file, mode, timeout);
                let (language, category, name) = model::identify(file);
                let exec = model::TestExecution {
                    path: file.clone(),
                    language,
                    category,
                    name,
                    result: outcome.result,
                    message: outcome.message,
                    duration_ms: outcome.duration_ms,
                };
                note(&exec);
                exec
            })
            .collect(),
        None => pool::run_all(
            &vybex,
            &files,
            threads,
            std::time::Duration::from_secs(timeout),
            note,
        ),
    };
    bar.finish_and_clear();

    let mut report = model::TestReport::new(under);
    for exec in executions {
        report.add_execution(exec);
    }
    report.duration_secs = started.elapsed().as_secs();

    report::print_console(&report);

    // Timestamped JSON, then diff against the newest earlier run of the same
    // runtime — a pass count on its own can't tell a fix from a regression.
    let stamp = chrono::Local::now().format("%Y%m%d_%H%M%S");
    let out = results_dir.join(format!("run_{stamp}.json"));
    if let Some((prev_path, prev)) = report::latest_previous(&results_dir, &report.runtime, &out) {
        report::compare(&prev, &report).print();
        eprintln!("[testrunner] previous: {}", prev_path.display());
    }
    report::save_json(&report, &out)?;
    println!("Report: {}", out.display());
    println!(
        "{}/{} passed ({:.1}%) in {:.1}s — {:.0} tests/s across {threads} workers",
        report.passed,
        report.total,
        report.pass_rate(),
        started.elapsed().as_secs_f64(),
        report.total as f64 / started.elapsed().as_secs_f64().max(0.001),
    );

    if report.failed == 0 && report.errors == 0 { Ok(()) } else { std::process::exit(1) }
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
