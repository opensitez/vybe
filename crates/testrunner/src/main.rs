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
mod style;
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
        Some("dashboard") => cmd_dashboard(&args[1..]),
        _ => {
            eprintln!(
                "vybe testrunner\n\n\
                 Usage:\n  \
                 testrunner extract <rust-test-file>... [--out DIR] [--lang NAME]\n  \
                 testrunner run <path>... [options]\n  \
                 testrunner summary <target>... [--results DIR]\n  \
                 testrunner dashboard [--results DIR] [--sort COL] [--desc]\n\n\
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
                 `extract --out DIR` writes DIR/<lang>/<category>/<test>. Pass\n\
                 the PARENT (`--out tests`), not the language directory — \n\
                 `--out tests/cobol` writes tests/cobol/cobol/… and merges a\n\
                 second copy of every category into the real corpus.\n\n\
                 `summary` reads a log written by `run --save` and groups the\n\
                 failures by category, worst first. `dashboard` totals EVERY\n\
                 saved log, one row per language.\n\n\
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
            skip_next = !matches!(
                arg.as_str(),
                "--verbose" | "--cold" | "--progress" | "--json" | "--save"
            );
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
    let mut compile_only = 0usize;
    let mut unpairable: Vec<(String, String)> = Vec::new();
    let mut empty: Vec<String> = Vec::new();

    for input in files {
        let path = Path::new(input);
        let lang = flag(args, "--lang")
            .map(str::to_string)
            .or_else(|| language_of(path))
            .with_context(|| format!("cannot tell the language of {input}"))?;
        let ext = emit::extension(&lang)
            .with_context(|| format!("no source extension known for `{lang}`"))?;

        let text = std::fs::read_to_string(path).with_context(|| format!("reading {input}"))?;
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
        // A module that yields NOTHING is the one failure extraction could not
        // report. An unlisted macro name, a helper reached by a path, a program
        // built with `format!` — each looks identical to "this file has no
        // tests", and the gap only ever showed as a count against the cargo
        // log: 8 C modules, 31 COBOL, 11 C#, 9 Pascal, 7 PHP.
        if cases.is_empty() {
            // At COLUMN ZERO only. `helpers.rs` carries an indented `#[test]`
            // inside its `macro_rules!` body — that is the DEFINITION of a
            // test, not a test, and cargo runs none of them.
            let declares_tests = text.lines().any(|l| {
                l.starts_with("#[test]")
                    || (l.contains("! {")
                        && !l.starts_with(char::is_whitespace)
                        && !l.starts_with("macro_rules!"))
            });
            if declares_tests {
                empty.push(input.to_string());
            }
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
        // `cobc` rejects a base name longer than 31 characters ("invalid file
        // base name — length exceeds maximum"), which would leave those cases
        // untestable against the reference compiler. Measured: 31 passes, 32
        // does not. The directory is not counted.
        let max_name = if lang == "cobol" { 31 } else { usize::MAX };
        let mut used_names: std::collections::HashSet<String> = Default::default();
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
                "cobol" => {
                    let e = emit::cobol::emit(case, input, &slug, &harness);
                    (e.text, e.pairing)
                }
                "fortran" => {
                    let e = emit::fortran::emit(case, input, &slug, &harness);
                    (e.text, e.pairing)
                }
                "dart" => {
                    let e = emit::dart::emit(case, input, &slug, &harness);
                    (e.text, e.pairing)
                }
                "pascal" => {
                    let e = emit::pascal::emit(case, input, &slug, &harness);
                    (e.text, e.pairing)
                }
                "c" => {
                    let e = emit::c::emit(case, input, &slug, &harness);
                    (e.text, e.pairing)
                }
                "lua" => {
                    let e = emit::lua::emit(case, input, &slug, &harness);
                    (e.text, e.pairing)
                }
                other => anyhow::bail!("no emitter for `{other}` yet"),
            };
            if let emit::go::Pairing::Unpairable(reason) = &pairing {
                unpairable.push((slug.clone(), reason.clone()));
            }
            // A compile case has no output to pair, so it is neither paired nor
            // unpairable — counting it as "paired 1:1 into assertions" would
            // overstate what the corpus checks.
            if case.expected.is_none() {
                compile_only += 1;
            }
            let file_name = short_unique_name(&case.name, max_name, &mut used_names);
            std::fs::write(dir.join(format!("{file_name}.{file_ext}")), &text)?;
            total += 1;
        }
        println!("{input} → {} case(s) in {}", cases.len(), dir.display());
    }

    println!("\nextracted {total} file(s)");
    if !empty.is_empty() {
        println!(
            "{} module(s) HAVE tests but yielded none — an unread shape:",
            empty.len()
        );
        for input in &empty {
            println!("  {input}");
        }
    }
    if compile_only > 0 {
        println!(
            "{compile_only} compile-mode case(s): the source must be ACCEPTED, \
             nothing is executed"
        );
    }
    if unpairable.is_empty() {
        println!(
            "{} case(s) paired 1:1 into assertions",
            total - compile_only
        );
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

// ── dashboard ───────────────────────────────────────────────────────────────

/// `testrunner dashboard` — one row per language across every saved log.
///
/// The `count_results.sh` equivalent for `run --save` output, with three
/// differences that matter:
///
/// * It counts VERDICTS, not lines containing "FAIL". A grep also matches the
///   `failures:` block, the `---- slug ----` headers and any harness
///   `FAIL: want [...]` the program printed, so the same run reads worse than
///   it is.
/// * It groups by the LANGUAGE in each slug, not the filename. Three saved
///   php logs are one `php` row; `${base%%.*}` would have labelled all three
///   `tests`.
/// * `TIMEOUT` is a verdict now, so timeouts are counted rather than inferred
///   from a "has been running" line that no longer exists.
fn cmd_dashboard(args: &[String]) -> Result<()> {
    // `watch testrunner dashboard` loses the colour, because `watch` is not a
    // terminal and `style` correctly refuses to paint a pipe. `--watch` redraws
    // from inside the process instead, where stdout IS the terminal.
    if !args.iter().any(|a| a == "--watch") {
        return render_dashboard(args);
    }
    let secs: u64 = flag(args, "--interval")
        .and_then(|v| v.parse().ok())
        .unwrap_or(2);
    loop {
        // Home + erase-below, not erase-all: the screen is overwritten in place
        // so a redraw does not flash.
        print!("\x1b[H\x1b[J");
        render_dashboard(args)?;
        println!(
            "\n{}",
            style::grey(&format!("refreshing every {secs}s — ^C to stop"))
        );
        use std::io::Write;
        let _ = std::io::stdout().flush();
        std::thread::sleep(std::time::Duration::from_secs(secs));
    }
}

fn render_dashboard(args: &[String]) -> Result<()> {
    let results_dir = PathBuf::from(flag(args, "--results").unwrap_or("results/testrunner"));
    let saved_dir = results_dir.join("saved");
    let sort_by = flag(args, "--sort")
        .unwrap_or("percent")
        .to_ascii_lowercase();
    let desc = args.iter().any(|a| a == "--desc");

    #[derive(Default)]
    struct Row {
        ok: usize,
        failed: usize,
        timeout: usize,
        files: usize,
        /// Tests the in-flight log said it would run, when one is running.
        expected: usize,
        running: bool,
        /// The owning process is gone but the log never finished.
        interrupted: bool,
        /// Seconds since the newest of this suite's logs was written.
        age: Option<u64>,
    }
    let mut rows: std::collections::BTreeMap<String, Row> = Default::default();

    // The suites are the directories under `tests/` — the same list that gives
    // a never-run suite its row further down. Each has exactly ONE log,
    // `saved/tests.<suite>.txt`, and the table reads that and nothing else.
    //
    // Reading every `.txt` in the directory counted a test once per file it
    // appeared in, and that directory accumulates files that are not the
    // suite's log: per-category runs (`tests.cobol.category_data_editing.txt`)
    // and duplicates of a whole suite (`fortran.txt` beside
    // `tests.fortran.txt`). Summed, an eight-day-old fortran log and today's
    // reported 11920 tests for a 5954-test suite at a blended 73% — neither
    // run's number — and `done` ran past 100%, because `expected` came from the
    // one log still in flight while ok/fail came from both.
    let extracted = extracted_suites();
    let logs: Vec<PathBuf> = extracted
        .iter()
        .map(|suite| saved_dir.join(format!("tests.{suite}.txt")))
        .filter(|log| log.is_file())
        .collect();
    anyhow::ensure!(
        !logs.is_empty(),
        "no suite logs in {} — run `testrunner run tests/<suite> --save` first",
        saved_dir.display()
    );

    let now = std::time::SystemTime::now();
    for log in &logs {
        let text = std::fs::read_to_string(log)?;
        // A run in flight has written its `running N tests` header and some
        // verdicts, but not the closing `test result:` line — `--save` streams
        // and only appends the tail at the end. No lockfile needed, and two
        // testrunners writing different suites are distinguished for free.
        let finished = text.lines().any(|l| l.starts_with("test result:"));
        let announced: usize = text
            .lines()
            .find_map(|l| {
                l.strip_prefix("running ")?
                    .split_whitespace()
                    .next()?
                    .parse()
                    .ok()
            })
            .unwrap_or(0);
        // Ask the owning process, if a sidecar names one.
        let owner_pid: Option<u32> = std::fs::read_to_string(run_marker(log))
            .ok()
            .and_then(|t| t.trim().parse().ok());
        let live = owner_pid.is_some_and(pid_alive);
        let orphaned = owner_pid.is_some() && !live;
        let age = std::fs::metadata(log)
            .and_then(|m| m.modified())
            .ok()
            .and_then(|t| now.duration_since(t).ok())
            .map(|d| d.as_secs());
        // The log's NAME already carries its target, and that is the only way
        // a log with no verdicts YET can appear at all: with `--suites`, the
        // savers for later waves exist, hold a live PID and are empty for as
        // long as the earlier waves take. Keyed off verdicts alone those suites
        // vanished from the table and reappeared under "never saved" — advice
        // to start a run that was already running.
        let mut seen_langs: std::collections::BTreeSet<String> = Default::default();
        if let Some(named) = lang_from_log(log) {
            seen_langs.insert(named);
        }
        for line in text.lines() {
            let Some(rest) = line.strip_prefix("test ") else {
                continue;
            };
            let Some((slug, verdict)) = rest.rsplit_once(" ... ") else {
                continue;
            };
            let lang = slug.split('/').next().unwrap_or(slug).to_string();
            let row = rows.entry(lang.clone()).or_default();
            match verdict {
                "ok" => row.ok += 1,
                "TIMEOUT" => row.timeout += 1,
                "FAILED" => row.failed += 1,
                _ => continue,
            }
            seen_langs.insert(lang);
        }
        for lang in seen_langs {
            let row = rows.entry(lang).or_default();
            row.files += 1;
            if !finished {
                row.expected += announced;
                if live {
                    row.running = true;
                } else {
                    row.interrupted = true;
                }
            } else if orphaned {
                // Finished writing but the marker survived — the process died
                // between the last write and its own cleanup.
                row.interrupted = true;
            }
            row.age = match (row.age, age) {
                (Some(a), Some(b)) => Some(a.min(b)),
                (a, b) => a.or(b),
            };
        }
    }

    let pct = |r: &Row| -> f64 {
        let total = r.ok + r.failed + r.timeout;
        if total == 0 {
            0.0
        } else {
            100.0 * r.ok as f64 / total as f64
        }
    };

    let mut ordered: Vec<(&String, &Row)> = rows.iter().collect();
    ordered.sort_by(|a, b| {
        let o = match sort_by.as_str() {
            "name" | "suite" => a.0.cmp(b.0),
            "ok" | "pass" => a.1.ok.cmp(&b.1.ok),
            "fail" | "failed" => a.1.failed.cmp(&b.1.failed),
            "timeout" => a.1.timeout.cmp(&b.1.timeout),
            "total" => {
                (a.1.ok + a.1.failed + a.1.timeout).cmp(&(b.1.ok + b.1.failed + b.1.timeout))
            }
            _ => pct(a.1)
                .partial_cmp(&pct(b.1))
                .unwrap_or(std::cmp::Ordering::Equal),
        };
        if desc { o.reverse() } else { o }
    });

    // `run_lang_tests.py`'s columns and vocabulary: a state icon on the left,
    // then the suite name BOLD and tinted by state. There is no `state` column
    // — the icon and the colour already say it, and `done` says how much of the
    // suite the row is speaking for. One format string serves the header and
    // every row, so the two cannot drift apart.
    //
    // EVERY CELL IS PADDED BEFORE IT IS COLOURED, and the macro's slots are
    // bare `{}`. A width applies to a string's LENGTH, and an escape sequence
    // is ~9 bytes of it, so `{:>6}` handed an already-coloured cell pads
    // nothing — the table lines up in a pipe (colour off) and comes apart on a
    // terminal, which is the one place anybody reads it.
    macro_rules! row {
        ($icon:expr, $name:expr, $pct:expr, $ok:expr, $fail:expr, $to:expr, $total:expr,
         $done:expr, $tail:expr) => {
            // Trimmed: a row with no tail would otherwise end in two spaces,
            // and trailing whitespace in a log is noise a diff will find.
            println!(
                "{}",
                format!(
                    "{} {} {} {} {} {} {} {}  {}",
                    $icon, $name, $pct, $ok, $fail, $to, $total, $done, $tail
                )
                .trim_end()
            )
        };
    }
    let head = |t: &str, w: usize| style::bold(&format!("{t:>w$}"));
    // A suite that was extracted but never saved still gets a row — `extracted`
    // is read at the top now, because it is also what chooses the logs.
    // The name column is as wide as the widest name it carries. A fixed 8 fit
    // every suite until `powershell` (10) — an over-long name does not truncate,
    // it pushes the rest of ITS row right, so one suite knocked every later
    // column out of line. The `.max(7)` keeps the table identical to before
    // whenever nothing is wider than the old fit.
    let name_w = rows
        .keys()
        .chain(extracted.iter())
        .map(|s| s.chars().count())
        .max()
        .unwrap_or(0)
        .max(7)
        + 1;
    // The columns through `done` are 42 wide plus the name; `updated` is free
    // text and is not ruled.
    let rule = style::grey(&"─".repeat(42 + name_w));
    row!(
        " ",
        style::bold(&format!("{:<name_w$}", "suite")),
        head("%ok", 6),
        head("ok", 6),
        head("fail", 6),
        head("t/o", 4),
        head("total", 7),
        head("done", 5),
        style::bold("updated")
    );
    println!("{rule}");

    let (mut t_ok, mut t_fail, mut t_to, mut t_want) = (0usize, 0usize, 0usize, 0usize);
    let dash = |w: usize| style::grey(&format!("{:>w$}", "—"));
    // A suite whose saver exists but has produced no verdict yet has nothing to
    // rank, so it sorts below the ones that do — with the never-saved rows,
    // which look the same because they mean the same thing: no data.
    let (measured, queued): (Vec<_>, Vec<_>) = ordered
        .iter()
        .partition(|(_, r)| r.ok + r.failed + r.timeout > 0);
    for (lang, r) in &measured {
        let total = r.ok + r.failed + r.timeout;
        // A finished log never announced an expectation, and it does not need
        // to: what it ran IS the suite.
        let want = if r.expected > 0 { r.expected } else { total };
        // GREEN MEANS LIVE. A suite that has finished is not tinted at all —
        // `run_lang_tests.py` leaves `done` uncoloured for the same reason:
        // colour on the name answers "is this moving?", and a finished suite
        // that stayed green answered it wrongly. How it SCORED is the `%ok`
        // column's job, and it has its own scale.
        let (icon, paint): (&str, fn(&str) -> String) = if r.running {
            ("▶", style::green)
        } else if r.interrupted {
            ("✖", style::red)
        } else if r.failed + r.timeout > 0 {
            ("✗", style::yellow)
        } else {
            ("✓", style::plain)
        };
        // Score, not liveness: ≥90% green, 80–90% orange, below that grey.
        let score: fn(&str) -> String = match pct(r) {
            p if p >= 90.0 => style::green,
            p if p >= 80.0 => style::orange,
            _ => style::grey,
        };
        let done = format!(
            "{:>5}",
            format!(
                "{:.0}%",
                if want == 0 {
                    0.0
                } else {
                    100.0 * total as f64 / want as f64
                }
            )
        );
        row!(
            paint(icon),
            style::bold(&paint(&format!("{lang:<name_w$}"))),
            score(&format!("{:>6}", format!("{:.1}%", pct(r)))),
            format!("{:>6}", r.ok),
            format!("{:>6}", r.failed),
            format!("{:>4}", r.timeout),
            format!("{total:>7}"),
            // A partial row is the one whose `done` matters, so it is the one
            // that gets the colour.
            if total == want {
                style::grey(&done)
            } else {
                paint(&done)
            },
            style::grey(&r.age.map(human_age).unwrap_or_else(|| "-".into()))
        );
        t_ok += r.ok;
        t_fail += r.failed;
        t_to += r.timeout;
        t_want += want;
    }
    // Its saver exists and its PID is alive, but its wave has not come up yet.
    // Zeros would read as "everything failing"; dashes read as what it is.
    for (lang, _) in &queued {
        row!(
            style::grey("·"),
            style::grey(&format!("{lang:<name_w$}")),
            dash(6),
            dash(6),
            dash(6),
            dash(4),
            dash(7),
            dash(5),
            style::grey("queued")
        );
    }

    // A suite that was extracted but never saved gets a ROW, not a sentence
    // trailing the table. `run_lang_tests.py` calls this state `queued` and
    // draws it grey with a `·`; here it means "no data", and a dashboard that
    // showed only what was measured would hide what was not. (`extracted` is
    // read above the header — it feeds the name column's width.)
    for lang in extracted.iter().filter(|l| !rows.contains_key(*l)) {
        row!(
            style::grey("·"),
            style::grey(&format!("{lang:<name_w$}")),
            dash(6),
            dash(6),
            dash(6),
            dash(4),
            dash(7),
            dash(5),
            // No tail. The grey `·` and the dashes already say there is no
            // data; spelling it out re-invents the status column.
            ""
        );
    }

    let grand = t_ok + t_fail + t_to;
    let overall = if grand == 0 {
        0.0
    } else {
        100.0 * t_ok as f64 / grand as f64
    };
    let live = ordered.iter().filter(|(_, r)| r.running).count();
    let stale = ordered.iter().filter(|(_, r)| r.interrupted).count();
    let bold_at = |t: String, w: usize| style::bold(&format!("{t:>w$}"));
    println!("{rule}");
    row!(
        " ",
        style::bold(&format!("{:<name_w$}", "TOTAL")),
        bold_at(format!("{overall:.1}%"), 6),
        bold_at(t_ok.to_string(), 6),
        bold_at(t_fail.to_string(), 6),
        bold_at(t_to.to_string(), 4),
        bold_at(grand.to_string(), 7),
        bold_at(
            format!(
                "{:.0}%",
                if t_want == 0 {
                    0.0
                } else {
                    100.0 * grand as f64 / t_want as f64
                }
            ),
            5
        ),
        ""
    );
    // A status line of its own, not a cell on the TOTAL row: a suite still
    // running, or one whose process died mid-log, makes its own percentage a
    // partial population, and that is a sentence, not a column.
    if live > 0 {
        println!(
            "{}",
            style::green(&format!("▶ {live} suite(s) still running"))
        );
    }
    if stale > 0 {
        println!(
            "{}",
            style::red(&format!(
                "✖ {stale} suite(s) interrupted — rerun before trusting those rows"
            ))
        );
    }
    Ok(())
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
    anyhow::ensure!(
        !targets.is_empty(),
        "no target given (e.g. `testrunner summary php`)"
    );

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
    let canonical = if dotted == "tests" || dotted.starts_with("tests.") {
        dotted.clone()
    } else {
        format!("tests.{dotted}")
    };
    let legacy = dotted.strip_prefix("tests.").unwrap_or(&dotted);
    let exact = [
        PathBuf::from(target),
        saved_dir.join(format!("{canonical}.txt")),
        saved_dir.join(format!("{legacy}.txt")),
        saved_dir.join(&canonical),
        saved_dir.join(legacy),
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
                && p.file_name().and_then(|n| n.to_str()).is_some_and(|n| {
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
        let Some(rest) = line.strip_prefix("test ") else {
            continue;
        };
        let Some((slug, verdict)) = rest.rsplit_once(" ... ") else {
            continue;
        };
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
        if timed_out > 0 {
            format!(" — {timed_out} timed out")
        } else {
            String::new()
        },
    );
    Ok(())
}

/// Say which binary is being driven, and say it loudly when it is a debug one.
///
/// A debug `vybex` is roughly an order of magnitude slower to COMPILE a
/// program, and the deadline is per test, so the difference does not show up as
/// "slower" — it shows up as TIMEOUT on whichever tests happen to sit near the
/// line. Measured 2026-08-04: all 148 Pascal timeouts were tests that inject
/// the Generics.Collections prelude, 27–56s each under `target/debug/vybex` and
/// ~5s under release. Zero of them time out on the release binary. A verdict
/// that flips with the profile is worse than a slow run, so name the binary.
/// When this run started, and when the binary it drives was built.
///
/// A saved log is otherwise undatable from the inside. Its mtime tells you when
/// the run FINISHED, and nothing at all tells you which `vybex` produced it —
/// so a log written minutes before a rebuild is indistinguishable from one
/// written after, and gets read as evidence for code it never executed.
/// `testrunner` deliberately links nothing from Vybe (see `Cargo.toml`), so it
/// is not even rebuilt when the compiler changes and its own mtime says
/// nothing. The binary's mtime is the honest answer to "is this result current".
fn provenance(vybex: &std::path::Path) -> String {
    let built = std::fs::metadata(vybex)
        .and_then(|m| m.modified())
        .map(|t| {
            chrono::DateTime::<chrono::Local>::from(t)
                .format("%Y-%m-%d %H:%M:%S")
                .to_string()
        })
        .unwrap_or_else(|_| "unknown".into());
    format!(
        "started {} · {} built {built}",
        chrono::Local::now().format("%Y-%m-%d %H:%M:%S"),
        vybex.display(),
    )
}

fn warn_if_debug_binary(vybex: &std::path::Path) {
    let looks_debug = vybex.components().any(|c| c.as_os_str() == "debug");
    let release = std::path::Path::new("target/release/vybex");
    if looks_debug && release.exists() {
        eprintln!(
            "note: driving {} — a DEBUG build. Compile-bound tests can TIME OUT \
             here and pass under --vybex target/release/vybex.",
            vybex.display()
        );
    }
}

// ── run ─────────────────────────────────────────────────────────────────────

fn cmd_run(args: &[String]) -> Result<()> {
    let vybex = PathBuf::from(flag(args, "--vybex").unwrap_or("target/debug/vybex"));
    warn_if_debug_binary(&vybex);
    let runtime: Option<Vec<String>> =
        flag(args, "--runtime").map(|cmd| cmd.split_whitespace().map(str::to_string).collect());
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
    let timeout: u64 = flag(args, "--timeout")
        .and_then(|n| n.parse().ok())
        .unwrap_or(60);
    let results_dir = PathBuf::from(flag(args, "--results").unwrap_or("results/testrunner"));
    // How many SUITES are in flight at once — `run_lang_tests.py`'s `jobs`, not
    // `-j`, which is worker threads. Several suites given at once used to be
    // flattened into one pool, so nothing finished until nearly everything did:
    // eleven `--save` logs all sat half-written, and the dashboard could only
    // say "running" about all of them. The cost is a drain-down at each wave
    // boundary, where one hung test holds the pool open for up to the full
    // timeout; raise the number to trade per-suite reporting back for it.
    let suites_at_once: usize = flag(args, "--suites")
        .and_then(|n| n.parse().ok())
        .unwrap_or(3)
        .max(1);
    if let Some(jobs) = flag(args, "-j").and_then(|n| n.parse().ok()) {
        rayon::ThreadPoolBuilder::new()
            .num_threads(jobs)
            .build_global()
            .ok();
    }

    let root = PathBuf::from(flag(args, "--tests").unwrap_or("tests"));
    let targets = positionals(args);
    anyhow::ensure!(!targets.is_empty(), "no suites or paths given");
    let mut roots = Vec::new();
    for target in &targets {
        let resolved = suites::resolve(target, &root).with_context(|| {
            format!(
                "no such suite or path: `{target}` (looked for it and for {}/{target})",
                root.display()
            )
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
        .unwrap_or_else(|| {
            std::thread::available_parallelism()
                .map(|n| n.get() + 2)
                .unwrap_or(8)
        });
    let under = match &runtime {
        Some(cmd) => cmd.join(" "),
        None => vybex.display().to_string(),
    };
    let waves: Vec<(Vec<usize>, Vec<PathBuf>)> = (0..roots.len())
        .collect::<Vec<_>>()
        .chunks(suites_at_once)
        .map(|group| {
            let group = group.to_vec();
            // ROUND-ROBIN across the wave's targets, not sorted path order.
            // The wave is one flat pool, and a pool fed in path order drains
            // `tests/java/...` entirely before it reaches `tests/wast/...` —
            // three suites "at a time" that still ran one at a time, with two
            // logs sitting empty while the first filled. Interleaving means
            // every suite in the wave progresses together.
            let mut lanes: Vec<Vec<PathBuf>> = group
                .iter()
                .map(|i| {
                    files
                        .iter()
                        .filter(|f| owner.get(*f) == Some(i))
                        .cloned()
                        .collect::<Vec<_>>()
                })
                .collect();
            let mut wave = Vec::new();
            let deepest = lanes.iter().map(Vec::len).max().unwrap_or(0);
            for n in 0..deepest {
                for lane in &mut lanes {
                    if let Some(file) = lane.get(n) {
                        wave.push(file.clone());
                    }
                }
            }
            (group, wave)
        })
        .collect();
    eprintln!(
        "[testrunner] {} test(s){} · {threads} {} workers · {timeout}s timeout · under `{under}`",
        files.len(),
        if roots.len() > 1 {
            format!(" · {} suite(s), {suites_at_once} at a time", roots.len())
        } else {
            String::new()
        },
        if runtime.is_some() || cold {
            "cold"
        } else {
            "warm"
        },
    );

    let mut counts: std::collections::BTreeMap<String, usize> = Default::default();
    for file in &files {
        *counts.entry(model::identify(file).0).or_default() += 1;
    }
    // The live table and the cargo-style per-test lines cannot share a stream:
    // the table redraws in place. Exactly one of them is on.
    let table = suites::Table::new(&counts, progress);
    let provenance = provenance(&vybex);
    if !progress {
        println!("\nrunning {} tests", files.len());
        println!("{provenance}");
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
    let savers: Vec<(PathBuf, std::sync::Mutex<Option<std::fs::File>>)> = if save {
        let mut out = Vec::new();
        for root in roots.iter() {
            let path = saved_log_path(&results_dir, root);
            if let Some(dir) = path.parent() {
                std::fs::create_dir_all(dir)?;
            }
            out.push((path, std::sync::Mutex::new(None)));
        }
        out
    } else {
        Vec::new()
    };
    // Each log is CREATED when its wave starts, not when the run starts.
    // Creating all of them up front truncated the logs of suites that would not
    // begin for hours, so a dashboard lost every earlier number the moment a
    // long multi-suite run began — and showed an empty file in its place.
    // Opened before the wave rather than at its end, and flushed per line, so
    // `tail -f` and a mid-run `grep` both see it fill.
    let open_saver = |i: usize| -> Result<()> {
        let Some((path, slot)) = savers.get(i) else {
            return Ok(());
        };
        let count = owner.values().filter(|&&o| o == i).count();
        let mut file =
            std::fs::File::create(path).with_context(|| format!("creating {}", path.display()))?;
        writeln!(file, "running {count} tests")?;
        // Second line, so a saved log carries its own date and the identity of
        // the binary that produced it — see [`provenance`].
        writeln!(file, "{provenance}")?;
        file.flush()?;
        // A sidecar naming the PID that owns this log. mtime cannot tell a live
        // run from an abandoned one — a stalled worker writes nothing for the
        // whole timeout, and a run killed a second ago still looks fresh. A PID
        // can be asked. Removed on the way out; if the process is killed it
        // stays behind and the dashboard sees a dead PID, which is exactly the
        // "interrupted" it should report.
        let _ = std::fs::write(run_marker(path), format!("{}\n", std::process::id()));
        *slot.lock().unwrap() = Some(file);
        Ok(())
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
            if let Ok(mut slot) = file.lock() {
                if let Some(file) = slot.as_mut() {
                    let _ = writeln!(file, "{line}");
                    let _ = file.flush();
                }
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

    let run_wave = |files: &[PathBuf]| -> Vec<model::TestExecution> {
        match &runtime {
            // A foreign runtime has no warm mode — one process per test is all
            // `go run` / `python3` / `node` offer.
            Some(cmd) => run::run_each(
                files,
                threads,
                |file, mode| {
                    run::run_foreign(&cmd[0], &cmd[1..], file, mode, timeout, &|secs| {
                        announce(file, secs)
                    })
                },
                &note,
            ),
            None if cold => run::run_each(
                files,
                threads,
                |file, mode| {
                    run::run_case(&vybex, file, mode, timeout, &|secs| announce(file, secs))
                },
                &note,
            ),
            None => pool::run_all(
                &vybex,
                files,
                threads,
                std::time::Duration::from_secs(timeout),
                &note,
                &announce,
            ),
        }
    };
    // One wave at a time, each wave being up to `--suites` targets run across
    // the full worker pool. With a single target there is one wave and this is
    // exactly the old behaviour.
    let mut executions: Vec<model::TestExecution> = Vec::with_capacity(files.len());
    for (group, wave) in &waves {
        // Each wave times itself. Elapsed-since-run-start made every suite
        // after the first overstate its duration in the very log a stats
        // script parses.
        let wave_started = std::time::Instant::now();
        for &i in group {
            open_saver(i)?;
        }
        let done = run_wave(wave);
        // Close each finished target's log HERE, not at the end of the whole
        // run. Its `test result:` line is what tells the dashboard the suite
        // completed; deferring it left a finished suite reading "running" for
        // as long as the remaining waves took.
        for &i in group {
            let Some((path, file)) = savers.get(i) else {
                continue;
            };
            let mut sub = model::TestReport::new(under.clone());
            for exec in done.iter().filter(|e| owner.get(&e.path) == Some(&i)) {
                sub.add_execution(exec.clone());
            }
            let mut slot = file.lock().unwrap();
            let Some(handle) = slot.as_mut() else {
                continue;
            };
            for line in report::cargo_tail(&sub, wave_started.elapsed().as_secs_f64(), false) {
                writeln!(handle, "{line}")?;
            }
            handle.flush()?;
            *slot = None;
            drop(slot);
            let _ = std::fs::remove_file(run_marker(path));
            eprintln!("[testrunner] saved: {}", path.display());
        }
        executions.extend(done);
    }
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

    if report.failed == 0 && report.errors == 0 {
        Ok(())
    } else {
        std::process::exit(1)
    }
}

/// Shorten a case name to `max` characters, keeping it unique within its
/// directory. Truncation alone collides (many names share a long prefix), so a
/// colliding name gets a numeric suffix inside the budget rather than beyond
/// it.
fn short_unique_name(
    name: &str,
    max: usize,
    used: &mut std::collections::HashSet<String>,
) -> String {
    let mut candidate: String = if name.chars().count() <= max {
        name.to_string()
    } else {
        name.chars().take(max).collect()
    };
    if used.insert(candidate.clone()) {
        return candidate;
    }
    for n in 2..1000u32 {
        let suffix = format!("_{n}");
        let keep = max.saturating_sub(suffix.chars().count());
        candidate = name.chars().take(keep).collect::<String>() + &suffix;
        if used.insert(candidate.clone()) {
            return candidate;
        }
    }
    candidate
}

/// `tests/php` → `results/testrunner/saved/tests.php.txt`.
///
/// The target with `/` turned into `.`, which is the shape the existing
/// `<lang>.tests.txt` stats scripts already read. **One file per target** —
/// `run tests/go tests/js --save` writes `tests.go.txt` and `tests.js.txt`,
/// not one merged log, because each is compared against its own suite.
/// The log for a suite, named after the RESOLVED suite path rather than the
/// spelling the caller typed.
///
/// `run go` and `run tests/go` are the same suite. Naming the log after the
/// argument gave them `go.txt` and `tests.go.txt` — two complete logs of the
/// same tests. Nothing downstream could tell them apart, because the dashboard
/// groups rows by the LANGUAGE in each slug (deliberately, so
/// `run tests/go/a tests/go/b` sums to one row). So the duplicate did not read
/// as a duplicate: it doubled ok/fail/total for that suite, and when one of the
/// pair was unfinished its announcement became the denominator for BOTH —
/// `done` at 191%.
fn saved_log_path(results_dir: &Path, resolved: &Path) -> PathBuf {
    // Relative to the working directory when possible, so an absolute target
    // does not spell its whole path into the filename.
    let rel = std::env::current_dir()
        .ok()
        .and_then(|cwd| resolved.strip_prefix(cwd).ok().map(Path::to_path_buf))
        .unwrap_or_else(|| resolved.to_path_buf());
    let name = rel.to_string_lossy().trim_matches('/').replace('/', ".");
    let name = if name.is_empty() {
        "run".to_string()
    } else {
        name
    };
    results_dir.join("saved").join(format!("{name}.txt"))
}

/// The suites, as the directories under `tests/` — the one list that says which
/// suites exist, so a suite with no log yet still gets a row and a suite's log
/// can be looked up by name rather than found by globbing.
fn extracted_suites() -> std::collections::BTreeSet<String> {
    std::fs::read_dir("tests")
        .into_iter()
        .flatten()
        .flatten()
        .filter(|entry| entry.path().is_dir())
        .filter_map(|entry| entry.file_name().into_string().ok())
        .collect()
}

/// `tests.php.txt` → `php`, `tests.php.arrays.txt` → `php`, `go.txt` → `go`.
/// The saved name is the target with `/` → `.`, so the suite is the first
/// component that is not the `tests` root.
fn lang_from_log(log: &Path) -> Option<String> {
    let stem = log.file_stem()?.to_str()?;
    stem.split('.')
        .find(|p| !p.is_empty() && *p != "tests")
        .map(str::to_string)
}

/// `12s` / `4m` / `2h` — a raw second count is unreadable in a `watch` pane.
/// `tests.php.txt` → `.tests.php.txt.run`, the sidecar holding the owning PID.
fn run_marker(log: &Path) -> PathBuf {
    let name = log
        .file_name()
        .map(|n| n.to_string_lossy().into_owned())
        .unwrap_or_default();
    log.with_file_name(format!(".{name}.run"))
}

/// Is that process still alive? `kill -0` is the portable unix probe: it sends
/// no signal and only reports whether the process exists.
fn pid_alive(pid: u32) -> bool {
    std::process::Command::new("kill")
        .args(["-0", &pid.to_string()])
        .stdout(std::process::Stdio::null())
        .stderr(std::process::Stdio::null())
        .status()
        .map(|s| s.success())
        .unwrap_or(false)
}

fn human_age(secs: u64) -> String {
    match secs {
        0..=99 => format!("{secs}s ago"),
        100..=5399 => format!("{}m ago", secs / 60),
        _ => format!("{}h ago", secs / 3600),
    }
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
