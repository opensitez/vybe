//! Console + JSON reporting, and run-over-run comparison.
//!
//! Ported from `ecma/testecma/src/report.rs`. The comparison is the part that
//! earns its keep: a bare pass count tells you nothing, `280 → 285 (+5)` with
//! the names of what broke tells you whether to keep going or go back.

use crate::model::{TestReport, TestResult};
use std::fs;
use std::path::{Path, PathBuf};

pub fn save_json(report: &TestReport, path: &Path) -> anyhow::Result<()> {
    if let Some(dir) = path.parent() {
        fs::create_dir_all(dir)?;
    }
    fs::write(path, serde_json::to_string_pretty(report)?)?;
    Ok(())
}

pub fn load(path: &Path) -> anyhow::Result<TestReport> {
    Ok(serde_json::from_str(&fs::read_to_string(path)?)?)
}

/// The newest report in `dir` for the same runtime, excluding `skip`.
pub fn latest_previous(dir: &Path, runtime: &str, skip: &Path) -> Option<(PathBuf, TestReport)> {
    let mut candidates: Vec<PathBuf> = fs::read_dir(dir)
        .ok()?
        .flatten()
        .map(|e| e.path())
        .filter(|p| p != skip && p.extension().map(|e| e == "json").unwrap_or(false))
        .collect();
    candidates.sort();
    candidates.into_iter().rev().find_map(|p| match load(&p) {
        Ok(r) if r.runtime == runtime => Some((p, r)),
        _ => None,
    })
}

pub fn print_console(report: &TestReport) {
    let rule = "=".repeat(78);
    println!("\n{rule}");
    println!("VYBE TESTRUNNER REPORT");
    println!("{rule}");
    println!("Timestamp: {}", report.timestamp);
    println!("Runtime:   {}", report.runtime);
    println!("Duration:  {} seconds", report.duration_secs);
    println!();

    println!("SUMMARY");
    println!("  Total:    {} tests", report.total);
    println!("  Passed:   {} ({:.1}%)", report.passed, report.pass_rate());
    println!("  Failed:   {}", report.failed);
    if report.skipped > 0 {
        println!("  Skipped:  {}", report.skipped);
    }
    if report.errors > 0 {
        println!("  Errors:   {}", report.errors);
    }
    if report.timeouts > 0 {
        println!("  Timeouts: {} (included in failed)", report.timeouts);
    }
    println!();

    if report.by_language.len() > 1 {
        println!("BY LANGUAGE");
        let mut langs: Vec<_> = report.by_language.iter().collect();
        langs.sort_by(|a, b| b.1.total.cmp(&a.1.total));
        for (name, stats) in langs {
            println!(
                "  {:<20} {}/{} ({:.1}%)",
                name,
                stats.passed,
                stats.total,
                stats.pass_rate()
            );
        }
        println!();
    }

    // Worst categories first — where the next hour of work pays off, which a
    // flat "top by count" list buries.
    let mut weak: Vec<_> = report
        .by_category
        .iter()
        .filter(|(_, s)| s.passed < s.total)
        .collect();
    if !weak.is_empty() {
        weak.sort_by(|a, b| {
            (b.1.total - b.1.passed)
                .cmp(&(a.1.total - a.1.passed))
                .then(a.0.cmp(b.0))
        });
        println!("WEAKEST CATEGORIES (by failing count)");
        for (name, stats) in weak.iter().take(15) {
            println!(
                "  {:<44} {}/{} ({:.1}%)",
                name,
                stats.passed,
                stats.total,
                stats.pass_rate()
            );
        }
        println!();
    }

    let failures: Vec<_> = report
        .executions
        .iter()
        .filter(|e| e.result != TestResult::Pass && e.result != TestResult::Skip)
        .collect();
    if !failures.is_empty() {
        println!("FIRST FAILURES (showing up to 10 of {})", failures.len());
        for (i, exec) in failures.iter().take(10).enumerate() {
            println!("  {}. {}", i + 1, exec.slug());
            println!("     {}", exec.message);
        }
        println!();
    }

    println!("{rule}");
}

/// The per-suite summary table, in the shape `run_lang_tests.py` prints.
pub fn print_summary_table(report: &TestReport, elapsed_secs: f64) {
    println!("\n=== summary ===");
    println!(
        "  {:<8} {:>7} {:>7} {:>7} {:>5} {:>7} {:>6}",
        "suite", "%ok", "ok", "fail", "t/o", "total", "time"
    );

    let mut timeouts: std::collections::HashMap<&str, usize> = std::collections::HashMap::new();
    for exec in &report.executions {
        if exec.result == TestResult::Timeout {
            *timeouts.entry(exec.language.as_str()).or_default() += 1;
        }
    }

    let mut rows: Vec<_> = report.by_language.iter().collect();
    // Worst first — the suite needing attention is the one you read first.
    rows.sort_by(|a, b| {
        a.1.pass_rate()
            .partial_cmp(&b.1.pass_rate())
            .unwrap_or(std::cmp::Ordering::Equal)
            .then(a.0.cmp(b.0))
    });

    for (lang, stats) in rows {
        let tos = timeouts.get(lang.as_str()).copied().unwrap_or(0);
        println!(
            "  {:<8} {:>7} {:>7} {:>7} {:>5} {:>7} {:>5}s",
            lang,
            format!("{:.2}", stats.pass_rate()),
            stats.passed,
            stats.failed,
            tos,
            stats.total,
            elapsed_secs.round() as u64,
        );
    }

    let fail_pct = if report.total > 0 {
        100.0 * (report.failed + report.errors) as f64 / report.total as f64
    } else {
        0.0
    };
    let tos_total: usize = timeouts.values().sum();
    let tos_str = if tos_total > 0 {
        format!(" / {tos_total} t/o")
    } else {
        String::new()
    };
    println!(
        " total: {} ok / {} fail{tos_str} ({fail_pct:.1}% fail)   {} suite(s)",
        report.passed,
        report.failed + report.errors,
        report.by_language.len(),
    );
}

/// The cargo-shaped tail: the failures block, then the `test result:` line.
///
/// Returned as lines rather than printed so the same text can go to a terminal
/// **with** colour and to a `--save` file **without** it. Formatting it twice
/// would let the two drift, and the saved file is what a stats script parses.
pub fn cargo_tail(report: &TestReport, elapsed_secs: f64, color: bool) -> Vec<String> {
    let paint = |kind: TestResult, text: &str| -> String {
        if !color {
            return text.to_string();
        }
        match kind {
            TestResult::Pass => crate::style::green(text),
            TestResult::Timeout => crate::style::orange(text),
            _ => crate::style::red(text),
        }
    };

    let failures: Vec<_> = report
        .executions
        .iter()
        .filter(|e| e.result != TestResult::Pass && e.result != TestResult::Skip)
        .collect();

    let mut out = Vec::new();
    if !failures.is_empty() {
        out.push(String::new());
        out.push("failures:".into());
        out.push(String::new());
        for exec in &failures {
            out.push(paint(exec.result, &format!("---- {} ----", exec.slug())));
            out.push(exec.message.clone());
            out.push(String::new());
        }
        out.push("failures:".into());
        for exec in &failures {
            out.push(format!("    {}", exec.slug()));
        }
    }

    out.push(String::new());
    out.push(format!(
        "test result: {}. {} passed; {} failed; {} ignored; 0 measured; 0 filtered out; \
         finished in {elapsed_secs:.2}s",
        if failures.is_empty() {
            paint(TestResult::Pass, "ok")
        } else {
            paint(TestResult::Fail, "FAILED")
        },
        report.passed,
        report.failed + report.errors,
        report.skipped,
    ));
    out
}

/// The per-test line, in cargo's shape. `color` off for a saved file.
pub fn cargo_verdict_line(exec: &crate::model::TestExecution, color: bool) -> String {
    // cargo's vocabulary, plus TIMEOUT. A hang is NOT a failure — the test
    // produced no answer at all — and calling it `FAILED` made the two
    // indistinguishable in a log. The word carries that; colour reinforces it.
    let (word, kind) = match exec.result {
        TestResult::Pass => ("ok", TestResult::Pass),
        TestResult::Skip => ("ignored", TestResult::Skip),
        TestResult::Timeout => ("TIMEOUT", TestResult::Timeout),
        _ => ("FAILED", TestResult::Fail),
    };
    let word = if !color {
        word.to_string()
    } else {
        match kind {
            TestResult::Pass => crate::style::green(word),
            TestResult::Skip => crate::style::yellow(word),
            TestResult::Timeout => crate::style::orange(word),
            _ => crate::style::red(word),
        }
    };
    format!("test {} ... {word}", exec.slug())
}

pub struct Diff {
    pub prev_timestamp: String,
    /// How many tests the two runs have in common — the only population a
    /// regression or fix can be drawn from.
    pub shared: usize,
    pub prev_total: usize,
    pub curr_total: usize,
    pub prev_pass: usize,
    pub curr_pass: usize,
    pub prev_fail: usize,
    pub curr_fail: usize,
    /// Passing before, failing now — the only list that should stop a merge.
    pub regressions: Vec<String>,
    pub fixes: Vec<String>,
}

pub fn compare(prev: &TestReport, curr: &TestReport) -> Diff {
    // Only tests present in BOTH runs can have moved. Without this, running the
    // whole corpus after running one module reports every other test as a
    // regression — "absent last time" is not "was passing last time".
    let prev_seen: std::collections::HashSet<String> =
        prev.executions.iter().map(|e| e.slug()).collect();
    let curr_seen: std::collections::HashSet<String> =
        curr.executions.iter().map(|e| e.slug()).collect();

    let before: std::collections::HashSet<String> = prev.failing_slugs().into_iter().collect();
    let after: std::collections::HashSet<String> = curr.failing_slugs().into_iter().collect();

    let mut regressions: Vec<String> = after
        .iter()
        .filter(|s| prev_seen.contains(*s) && !before.contains(*s))
        .cloned()
        .collect();
    let mut fixes: Vec<String> = before
        .iter()
        .filter(|s| curr_seen.contains(*s) && !after.contains(*s))
        .cloned()
        .collect();
    regressions.sort();
    fixes.sort();

    Diff {
        prev_timestamp: prev.timestamp.clone(),
        shared: prev_seen.intersection(&curr_seen).count(),
        prev_total: prev.total,
        curr_total: curr.total,
        prev_pass: prev.passed,
        curr_pass: curr.passed,
        prev_fail: prev.failed,
        curr_fail: curr.failed,
        regressions,
        fixes,
    }
}

impl Diff {
    pub fn print(&self) {
        let rule = "=".repeat(78);
        println!("{rule}");
        println!("COMPARED TO {}", self.prev_timestamp);
        println!("{rule}");
        if self.prev_total != self.curr_total {
            println!(
                "Test set changed: {} → {} ({} in common; only those are compared)",
                self.prev_total, self.curr_total, self.shared
            );
        }
        println!(
            "Passes: {} → {} ({:+})",
            self.prev_pass,
            self.curr_pass,
            self.curr_pass as i64 - self.prev_pass as i64
        );
        println!(
            "Fails:  {} → {} ({:+})",
            self.prev_fail,
            self.curr_fail,
            self.curr_fail as i64 - self.prev_fail as i64
        );
        if !self.regressions.is_empty() {
            println!("\nREGRESSIONS ({}):", self.regressions.len());
            for slug in self.regressions.iter().take(25) {
                println!("  ✗ {slug}");
            }
            if self.regressions.len() > 25 {
                println!("  … {} more", self.regressions.len() - 25);
            }
        }
        if !self.fixes.is_empty() {
            println!("\nNEWLY PASSING ({}):", self.fixes.len());
            for slug in self.fixes.iter().take(25) {
                println!("  ✓ {slug}");
            }
            if self.fixes.len() > 25 {
                println!("  … {} more", self.fixes.len() - 25);
            }
        }
        println!("{rule}");
    }
}
