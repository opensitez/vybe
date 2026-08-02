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
mod style;
mod rustlit;
mod suites;

use anyhow::{Context, Result};
use std::io::Write;
use std::path::{Path, PathBuf};
use walkdir::WalkDir;

fn main() -> Result<()> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    match args.first().map(String::as_str) {
        Some("extract") => cmd_extract(&args[1..]),
        Some("run") => cmd_run(&args[1..]),
        Some("summary") => cmd_summary(&args[1..]),
        _ => {
            eprintln!(
                "vybe testrunner\n\n\
                 Usage:\n  \
                 testrunner extract <rust-test-file>... [--out DIR] [--lang NAME]\n  \
                 testrunner run <path>... [options]\n  \
                 testrunner summary <target>... [--results DIR]\n\n\
                 run options:\n  \
                 --vybex PATH    vybex binary to drive      (default target/debug/vybex)\n  \
                 --runtime CMD   run under something else   (\"go run\", \"node\", \"python3\")\n  \
                 --cold          one fresh process per test (default: warm workers)\n  \
                 -j N            parallel workers           (default: CPU count)\n  \
                 --timeout SECS  per-test deadline          (default 60)\n  \
                 --results DIR   report directory           (default results/testrunner)\n  \
                 --progress      live per-suite table       (default: cargo-test output)\n  \
                 --json          write the timestamped JSON report + regression diff\n  \
                 --save          write the plain test log to results/<dir>/saved/<target>.txt\n  \
                 --verbose       stream every failure as it happens (default: report only)\n\n\
                 `summary` reads a log written by `run --save` and groups the\n\
                 failures by category, worst first.\n\n\
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
            skip_next =
                !matches!(arg.as_str(), "--verbose" | "--cold" | "--progress" | "--json" | "--save");
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
        // BOTH shapes, merged. A module can carry a macro batch AND its own
        // `#[test] fn`s — `test_bcmath.rs` has a 12-entry `php_cases!` block
        // beside 34 test functions, and treating the two as mutually exclusive
        // silently dropped whichever came second.
        // One unreadable module must not abort the whole extraction — it did,
        // and silently left a partial corpus behind.
        let mut cases = match extract::cases_in_file(&text) {
            Ok(cases) => cases,
            Err(e) => {
                eprintln!("  {input}: macro parse stopped ({e})");
                Vec::new()
            }
        };
        let seen: std::collections::HashSet<String> =
            cases.iter().map(|c| c.name.clone()).collect();
        cases.extend(
            extract::test_fns_in_file(&text)
                .into_iter()
                .filter(|c| !seen.contains(&c.name)),
        );
        let seen: std::collections::HashSet<String> =
            cases.iter().map(|c| c.name.clone()).collect();
        cases.extend(
            extract::paren_macros_in_file(&text)
                .into_iter()
                .filter(|c| !seen.contains(&c.name)),
        );
        let seen: std::collections::HashSet<String> =
            cases.iter().map(|c| c.name.clone()).collect();
        cases.extend(
            extract::paren_macros_in_file(&text)
                .into_iter()
                .filter(|c| !seen.contains(&c.name)),
        );
        if cases.is_empty() {
            continue;
        }

        let category = path
            .file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or("misc")
            .trim_start_matches("test_")
            .trim_end_matches("_test")
            .trim_start_matches("js_")
            .to_string();
        let dir = out_root.join(&lang).join(&category);
        std::fs::create_dir_all(&dir)?;
        // Not every language needs an injected harness — wast asserts by
        // trapping, so its compile-mode cases need no helper at all.
        let harness = emit::harness_body(&lang).unwrap_or_default();

        for case in &cases {
            round_trip_check(case, input)?;

            let slug = format!("{lang}/{category}/{}", case.name);
            // A wast case picks its own extension: a script carrying
            // assertions must be `.wast` for a spec interpreter to run its
            // directives, while a bare module stays `.wat`.
            let mut file_ext = ext;
            let (text, pairing) = match lang.as_str() {
                "go" => {
                    let e = emit::go::emit(case, input, &slug, &harness);
                    (e.text, e.pairing)
                }
                "js" => {
                    let e = emit::js::emit(case, input, &slug, &harness);
                    (e.text, e.pairing)
                }
                "php" => {
                    let e = emit::php::emit(case, input, &slug, &harness);
                    (e.text, e.pairing)
                }
                "kotlin" => {
                    let e = emit::kotlin::emit(case, input, &slug, &harness);
                    (e.text, e.pairing)
                }
                "csharp" => {
                    let e = emit::csharp::emit(case, input, &slug, &harness);
                    (e.text, e.pairing)
                }
                "wast" => {
                    let e = emit::wast::emit(case, input, &slug, &harness);
                    file_ext = e.extension;
                    (e.text, e.pairing)
                }
                "vb" => {
                    let e = emit::vb::emit(case, input, &slug, &harness);
                    (e.text, e.pairing)
                }
                "python" => {
                    let e = emit::python::emit(case, input, &slug, &harness);
                    (e.text, e.pairing)
                }
                "ruby" => {
                    let e = emit::ruby::emit(case, input, &slug, &harness);
                    (e.text, e.pairing)
                }
                "java" => {
                    let e = emit::java::emit(case, input, &slug, &harness);
                    (e.text, e.pairing)
                }
                other => anyhow::bail!("no emitter for `{other}` yet"),
            };
            if let emit::go::Pairing::Unpairable(reason) = &pairing {
                unpairable.push((slug.clone(), reason.clone()));
            }
            std::fs::write(dir.join(format!("{}.{file_ext}", case.name)), &text)?;
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

// ── summary ─────────────────────────────────────────────────────────────────

/// `testrunner summary php` — group a saved log's failures by category.
///
/// Reads the plain log `run --save` wrote, so it works long after the run and
/// needs nothing but the file. Parsing the verdict properly is why this is not
/// a `grep FAIL` one-liner: `FAILED` also appears in the `test result:` line,
/// and `TIMEOUT` is a separate verdict a naive grep would miss entirely.
fn cmd_summary(args: &[String]) -> Result<()> {
    let results_dir = PathBuf::from(flag(args, "--results").unwrap_or("results/testrunner"));
    let saved_dir = results_dir.join("saved");
    let targets = positionals(args);
    anyhow::ensure!(!targets.is_empty(), "no target given (e.g. `testrunner summary php`)");

    for target in targets {
        let logs = resolve_saved(&saved_dir, target);
        if logs.is_empty() {
            eprintln!(
                "[testrunner] no saved log for `{target}` in {}",
                saved_dir.display()
            );
            let mut available: Vec<String> = std::fs::read_dir(&saved_dir)
                .into_iter()
                .flatten()
                .flatten()
                .filter_map(|e| e.file_name().into_string().ok())
                .collect();
            available.sort();
            if available.is_empty() {
                eprintln!("  nothing saved yet. Run one:");
            } else {
                eprintln!("  available: {}", available.join(", "));
                eprintln!("  or save that target first:");
            }
            eprintln!("      testrunner run tests/{target} --save");
            std::process::exit(1);
        }
        print_saved_summary(target, &logs)?;
    }
    Ok(())
}

/// Accept whatever the user has in hand: `php`, `tests/php`, `tests.php`,
/// `php/bcmath`, or a path to the log itself.
///
/// Falling back to a PREFIX match is what makes `summary php` useful: you far
/// more often save a handful of categories than a whole suite, and an exact
/// lookup then reported "no saved log for php" with three php logs sitting in
/// the directory.
fn resolve_saved(saved_dir: &Path, target: &str) -> Vec<PathBuf> {
    let dotted = target.trim_matches('/').replace('/', ".");
    let exact = [
        PathBuf::from(target),
        saved_dir.join(format!("{dotted}.txt")),
        saved_dir.join(format!("tests.{dotted}.txt")),
        saved_dir.join(&dotted),
    ];
    if let Some(hit) = exact.into_iter().find(|p| p.is_file()) {
        return vec![hit];
    }

    // `php` also matches `tests.php.arrays.txt` and `php.arrays.txt`. The
    // trailing dot keeps `php` from matching a hypothetical `phpunit.*`.
    let mut matches: Vec<PathBuf> = std::fs::read_dir(saved_dir)
        .into_iter()
        .flatten()
        .flatten()
        .map(|e| e.path())
        .filter(|p| {
            p.is_file()
                && p.file_name()
                    .and_then(|n| n.to_str())
                    .is_some_and(|n| {
                        n.starts_with(&format!("{dotted}."))
                            || n.starts_with(&format!("tests.{dotted}."))
                    })
        })
        .collect();
    matches.sort();
    matches
}

fn print_saved_summary(target: &str, paths: &[PathBuf]) -> Result<()> {
    let mut text = String::new();
    for path in paths {
        text.push_str(
            &std::fs::read_to_string(path)
                .with_context(|| format!("reading {}", path.display()))?,
        );
        text.push('\n');
    }

    let mut by_category: std::collections::BTreeMap<String, (usize, usize)> = Default::default();
    let (mut passed, mut failed, mut timed_out, mut ignored) = (0usize, 0usize, 0usize, 0usize);

    for line in text.lines() {
        // `test <slug> ... <verdict>` and nothing else. The `test result:`
        // summary line and the `---- slug ----` headers must not be counted.
        let Some(rest) = line.strip_prefix("test ") else { continue };
        let Some((slug, verdict)) = rest.rsplit_once(" ... ") else { continue };
        let (fails, tos) = match verdict {
            "ok" => {
                passed += 1;
                continue;
            }
            "ignored" => {
                ignored += 1;
                continue;
            }
            "TIMEOUT" => {
                timed_out += 1;
                (0, 1)
            }
            "FAILED" => {
                failed += 1;
                (1, 0)
            }
            _ => continue,
        };
        // The slug is `lang/category/name`; group by `lang/category`, which is
        // the unit you re-run and the unit a fix lands in.
        let mut parts = slug.split('/');
        let category = match (parts.next(), parts.next(), parts.next()) {
            (Some(lang), Some(cat), Some(_)) => format!("{lang}/{cat}"),
            _ => slug.to_string(),
        };
        let entry = by_category.entry(category).or_default();
        entry.0 += fails;
        entry.1 += tos;
    }

    if let [only] = paths {
        println!("== {}", only.display());
    } else {
        // Say which logs were merged — a total over an unstated set of files
        // is not a number anyone can act on.
        println!("== {target} — {} logs merged", paths.len());
        for path in paths {
            println!("   {}", path.display());
        }
    }
    let total = passed + failed + timed_out + ignored;
    if total == 0 {
        println!("   (no test lines found — is this a `run --save` log?)");
        return Ok(());
    }

    // Worst first: where the next hour of work pays off.
    let mut rows: Vec<_> = by_category.into_iter().collect();
    rows.sort_by(|a, b| (b.1.0 + b.1.1).cmp(&(a.1.0 + a.1.1)).then(a.0.cmp(&b.0)));

    // The timeout column only appears when there are timeouts — otherwise it
    // is a column of zeroes and the output stops matching the familiar
    // `count category` shape.
    let any_timeouts = timed_out > 0;
    if any_timeouts && !rows.is_empty() {
        println!("{:>6} {:>5}  {}", "fail", "t/o", "category");
    }
    for (category, (fails, tos)) in &rows {
        if any_timeouts {
            println!("{fails:>6} {tos:>5}  {category}");
        } else {
            println!("{fails:>6}  {category}");
        }
    }

    let bad = failed + timed_out;
    println!(
        "   {passed} passed, {bad} not passing ({:.1}%) of {total}{}",
        100.0 * bad as f64 / total as f64,
        if timed_out > 0 { format!(" — {timed_out} timed out") } else { String::new() },
    );
    Ok(())
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
    // Default output is cargo-test shaped — `test <slug> ... ok` per line, then
    // a failures list and a `test result:` summary. It greps, it diffs against
    // `<lang>.tests.txt`, and a tool reading it needs no terminal. `--progress`
    // swaps in the live per-suite table for a human watching a long run.
    let progress = args.iter().any(|a| a == "--progress");
    // The timestamped JSON is opt-in. Most runs are one category while
    // iterating, and writing a report per run filled `results/` with hundreds
    // of files nobody read. The run-over-run diff rides with it — without a
    // saved report there is no baseline for a later run to compare against.
    let json_report = args.iter().any(|a| a == "--json");
    // Plain-text copy of the test log, named after the target, for a stats
    // script to parse the way it parses `<lang>.tests.txt` today.
    let save = args.iter().any(|a| a == "--save");
    // 60s, matching the point at which cargo declares a test worth mentioning.
    let timeout: u64 = flag(args, "--timeout").and_then(|n| n.parse().ok()).unwrap_or(60);
    let results_dir = PathBuf::from(flag(args, "--results").unwrap_or("results/testrunner"));
    if let Some(jobs) = flag(args, "-j").and_then(|n| n.parse().ok()) {
        rayon::ThreadPoolBuilder::new().num_threads(jobs).build_global().ok();
    }

    let root = PathBuf::from(flag(args, "--tests").unwrap_or("tests"));
    let targets = positionals(args);
    anyhow::ensure!(!targets.is_empty(), "no suites or paths given");
    let mut roots = Vec::new();
    for target in &targets {
        let resolved = suites::resolve(target, &root).with_context(|| {
            format!("no such suite or path: `{target}` (looked for it and for {}/{target})", root.display())
        })?;
        roots.push(resolved);
    }
    // Which target each test came from, so `--save` can write one file per
    // target rather than pooling several suites into one log. First target
    // wins if two overlap.
    let mut files = Vec::new();
    let mut owner: std::collections::HashMap<PathBuf, usize> = Default::default();
    for (i, root) in roots.iter().enumerate() {
        for file in collect(std::slice::from_ref(root)) {
            if owner.contains_key(&file) {
                continue;
            }
            owner.insert(file.clone(), i);
            files.push(file);
        }
    }
    files.sort();
    anyhow::ensure!(!files.is_empty(), "no test files found");

    // Workers are subprocesses that block on I/O as well as compute, so a
    // couple more than the core count keeps them all fed. Measured on a
    // 10-core box: 10→97, 12→99, 16→92, 20→88 tests/s.
    let threads = flag(args, "-j")
        .and_then(|n| n.parse().ok())
        .unwrap_or_else(|| std::thread::available_parallelism().map(|n| n.get() + 2).unwrap_or(8));
    let under = match &runtime {
        Some(cmd) => cmd.join(" "),
        None => vybex.display().to_string(),
    };
    eprintln!(
        "[testrunner] {} test(s) · {threads} {} workers · {timeout}s timeout · under `{under}`",
        files.len(),
        if runtime.is_some() || cold { "cold" } else { "warm" },
    );

    let mut counts: std::collections::BTreeMap<String, usize> = Default::default();
    for file in &files {
        *counts.entry(model::identify(file).0).or_default() += 1;
    }
    // The live table and the cargo-style per-test lines cannot share a stream:
    // the table redraws in place. Exactly one of them is on.
    let table = suites::Table::new(&counts, progress);
    if !progress {
        println!("\nrunning {} tests", files.len());
    }

    let started = std::time::Instant::now();
    // A WARNING, and shaped like one. It is not a verdict: the test has not
    // failed, has not timed out, and may still pass. Writing it as
    // `test <slug> …` made it read as a result line, which is exactly the
    // confusion to avoid. Without it a hang is a silent gap for the whole
    // timeout and looks like the runner stalled — it has not, the other workers
    // keep draining and the stuck one is killed and replaced at the deadline.
    let announce = |file: &std::path::Path, secs: u64| {
        let (lang, cat, name) = model::identify(file);
        table.suspend(|| {
            println!(
                "{} {lang}/{cat}/{name} still running after {secs}s",
                style::yellow("warning:")
            );
        });
    };
    // Opened BEFORE the run and written line by line, not buffered to the end:
    // on a suite that takes minutes you want to watch or grep the log while it
    // is still filling.
    let savers: Vec<(PathBuf, std::sync::Mutex<std::fs::File>)> = if save {
        let mut out = Vec::new();
        for (i, target) in targets.iter().enumerate() {
            let path = saved_log_path(&results_dir, target);
            if let Some(dir) = path.parent() {
                std::fs::create_dir_all(dir)?;
            }
            let count = owner.values().filter(|&&o| o == i).count();
            let mut file = std::fs::File::create(&path)
                .with_context(|| format!("creating {}", path.display()))?;
            writeln!(file, "running {count} tests")?;
            out.push((path, std::sync::Mutex::new(file)));
        }
        out
    } else {
        Vec::new()
    };
    let note = |exec: &model::TestExecution| {
        table.record(exec);
        if !progress {
            // cargo prints each verdict as it lands, unordered under parallelism.
            // Exactly cargo's three verdicts. A timeout is a FAILURE, not a
            // fourth word: anything parsing `ok|FAILED|ignored` would skip a
            // `TIMEOUT` line and under-report. The reason survives in the
            // failures block ("timeout after 30s").
            println!("{}", report::cargo_verdict_line(exec, true));
        }
        if let Some((_, file)) = owner.get(&exec.path).and_then(|i| savers.get(*i)) {
            // Always uncoloured: the file exists to be parsed. Flushed per
            // line so `tail -f` and a mid-run `grep` both see it.
            let line = report::cargo_verdict_line(exec, false);
            if let Ok(mut file) = file.lock() {
                let _ = writeln!(file, "{line}");
                let _ = file.flush();
            }
        }
        // Quiet by default. A failure is a data point, not a message: the live
        // rows carry the counts, the console report names the first few, and
        // the JSON holds all of them. Streaming every one turned a 705-failure
        // run into 1,410 lines of scrollback.
        if verbose && exec.result != model::TestResult::Pass {
            table.suspend(|| {
                println!("FAIL {}", exec.slug());
                println!("     {}", exec.message);
                println!("     {}", exec.path.display());
            });
        }
    };

    let executions: Vec<model::TestExecution> = match &runtime {
        // A foreign runtime has no warm mode — one process per test is all
        // `go run` / `python3` / `node` offer.
        Some(cmd) => run::run_each(&files, threads, |file, mode| {
            run::run_foreign(&cmd[0], &cmd[1..], file, mode, timeout, &|secs| {
                announce(file, secs)
            })
        }, &note),
        None if cold => run::run_each(&files, threads, |file, mode| {
            run::run_case(&vybex, file, mode, timeout, &|secs| announce(file, secs))
        }, &note),
        None => pool::run_all(
            &vybex,
            &files,
            threads,
            std::time::Duration::from_secs(timeout),
            note,
            &announce,
        ),
    };
    table.finish();

    let mut report = model::TestReport::new(under.clone());
    for exec in executions {
        report.add_execution(exec);
    }
    report.duration_secs = started.elapsed().as_secs();

    let secs = started.elapsed().as_secs_f64();
    if progress {
        report::print_console(&report);
        report::print_summary_table(&report, secs);
    } else {
        for line in report::cargo_tail(&report, secs, true) {
            println!("{line}");
        }
    }

    // Each target's file gets ITS OWN failures block and `test result:` line,
    // counted over that target's tests only — the file is that target's log.
    for (i, (path, file)) in savers.iter().enumerate() {
        let mut sub = model::TestReport::new(under.clone());
        for exec in report.executions.iter().filter(|e| owner.get(&e.path) == Some(&i)) {
            sub.add_execution(exec.clone());
        }
        let mut file = file.lock().unwrap();
        for line in report::cargo_tail(&sub, secs, false) {
            writeln!(file, "{line}")?;
        }
        file.flush()?;
        eprintln!("[testrunner] saved: {}", path.display());
    }

    // Timestamped JSON, and the run-over-run diff that rides on it. Opt-in:
    // one report per run buried `results/` in files nobody opened.
    let stamp = chrono::Local::now().format("%Y%m%d_%H%M%S");
    let out = results_dir.join(format!("run_{stamp}.json"));
    let previous = json_report
        .then(|| report::latest_previous(&results_dir, &report.runtime, &out))
        .flatten();
    if json_report {
        report::save_json(&report, &out)?;
    }

    if progress {
        if let Some((prev_path, prev)) = &previous {
            report::compare(prev, &report).print();
            eprintln!("[testrunner] previous: {}", prev_path.display());
        }
        if json_report {
            println!("Report: {}", out.display());
        }
        println!(
            "{}/{} passed ({:.1}%) in {:.1}s — {:.0} tests/s across {threads} workers",
            report.passed,
            report.total,
            report.pass_rate(),
            started.elapsed().as_secs_f64(),
            report.total as f64 / started.elapsed().as_secs_f64().max(0.001),
        );
    } else {
        // cargo-style already ended with `test result:`. Keep the extras on
        // stderr so stdout stays exactly the greppable test log.
        if let Some((_, prev)) = &previous {
            let diff = report::compare(prev, &report);
            if !diff.regressions.is_empty() || !diff.fixes.is_empty() {
                eprintln!(
                    "[testrunner] vs previous: {} regression(s), {} newly passing",
                    diff.regressions.len(),
                    diff.fixes.len()
                );
            }
        }
        if report.timeouts > 0 {
            // cargo's summary line has no timeout field, so it goes to stderr
            // rather than being wedged into a format tools parse. They are
            // still counted in `failed` there, or the numbers would not add up
            // to the total.
            eprintln!(
                "[testrunner] {} test(s) timed out (counted in failed above)",
                report.timeouts
            );
        }
        if json_report {
            eprintln!("[testrunner] report: {}", out.display());
        }
    }

    if report.failed == 0 && report.errors == 0 { Ok(()) } else { std::process::exit(1) }
}

/// `tests/php` → `results/testrunner/saved/tests.php.txt`.
///
/// The target with `/` turned into `.`, which is the shape the existing
/// `<lang>.tests.txt` stats scripts already read. **One file per target** —
/// `run tests/go tests/js --save` writes `tests.go.txt` and `tests.js.txt`,
/// not one merged log, because each is compared against its own suite.
fn saved_log_path(results_dir: &Path, target: &str) -> PathBuf {
    let name = target.trim_matches('/').replace('/', ".");
    let name = if name.is_empty() { "run".to_string() } else { name };
    results_dir.join("saved").join(format!("{name}.txt"))
}

fn collect(roots: &[PathBuf]) -> Vec<PathBuf> {
    let mut files = Vec::new();
    for root in roots {
        let path = root.as_path();
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
