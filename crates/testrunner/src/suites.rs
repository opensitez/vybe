//! Target resolution and the live per-suite table.
//!
//! Mirrors `run_lang_tests.py`: name suites rather than paths
//! (`testrunner run go python`), and show one live row per suite instead of a
//! single undifferentiated bar, so a slow or failing language is visible while
//! the run is still going.

use crate::model::{TestExecution, TestResult};
use indicatif::{MultiProgress, ProgressBar, ProgressStyle};
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicUsize, Ordering};

/// Resolve one argument to a path. A bare language name means that language's
/// whole suite, so `run go` and `run tests/go` are the same thing.
pub fn resolve(arg: &str, root: &Path) -> Option<PathBuf> {
    // A bare word is a SUITE NAME first. Checking the literal path first meant
    // `run php` picked up the repo's `php/` directory (php-src, langspec) and
    // reported "no test files found" while `tests/php` sat there full.
    if !arg.contains('/') && !arg.contains('.') {
        let by_name = root.join(arg);
        if by_name.exists() {
            return Some(by_name);
        }
    }
    let direct = PathBuf::from(arg);
    if direct.exists() {
        return Some(direct);
    }
    let by_name = root.join(arg);
    by_name.exists().then_some(by_name)
}

pub struct Suite {
    pub bar: ProgressBar,
    pub ok: AtomicUsize,
    pub failed: AtomicUsize,
    pub timeouts: AtomicUsize,
}

/// One live row per language, plus a total row underneath.
pub struct Table {
    multi: MultiProgress,
    suites: BTreeMap<String, Suite>,
    total: ProgressBar,
}

impl Table {
    /// `visible` is decided BEFORE any bar is added. Adding first and hiding
    /// after leaves a flash of half-drawn bars on a terminal, because
    /// `MultiProgress` paints on `add`.
    pub fn new(counts: &BTreeMap<String, usize>, visible: bool) -> Self {
        let multi = MultiProgress::new();
        if !visible {
            multi.set_draw_target(indicatif::ProgressDrawTarget::hidden());
        }
        let style = ProgressStyle::with_template(
            " {prefix} {bar:22.cyan/blue} {pos:>6}/{len:<6} {msg}",
        )
        .unwrap_or_else(|_| ProgressStyle::default_bar())
        .progress_chars("##-");

        let mut suites = BTreeMap::new();
        for (lang, count) in counts {
            let bar = multi.add(ProgressBar::new(*count as u64));
            bar.set_style(style.clone());
            // Escapes are zero-width, so the column still lines up.
            bar.set_prefix(format!("\x1b[1m{lang:<9}\x1b[0m"));
            bar.set_message("·");
            suites.insert(
                lang.clone(),
                Suite {
                    bar,
                    ok: AtomicUsize::new(0),
                    failed: AtomicUsize::new(0),
                    timeouts: AtomicUsize::new(0),
                },
            );
        }

        let grand_total: usize = counts.values().sum();
        let total = multi.add(ProgressBar::new(grand_total as u64));
        total.set_style(
            ProgressStyle::with_template(" {prefix} {bar:22.green/blue} {pos:>6}/{len:<6} {msg}")
                .unwrap_or_else(|_| ProgressStyle::default_bar())
                .progress_chars("##-"),
        );
        total.set_prefix(format!("\x1b[1m{:<9}\x1b[0m", "total"));

        Table { multi, suites, total }
    }

    /// Fold one finished test into its suite's row.
    pub fn record(&self, exec: &TestExecution) {
        let Some(suite) = self.suites.get(&exec.language) else {
            return;
        };
        match exec.result {
            TestResult::Pass => {
                suite.ok.fetch_add(1, Ordering::Relaxed);
            }
            TestResult::Timeout => {
                suite.failed.fetch_add(1, Ordering::Relaxed);
                suite.timeouts.fetch_add(1, Ordering::Relaxed);
            }
            _ => {
                suite.failed.fetch_add(1, Ordering::Relaxed);
            }
        }

        let ok = suite.ok.load(Ordering::Relaxed);
        let bad = suite.failed.load(Ordering::Relaxed);
        let tos = suite.timeouts.load(Ordering::Relaxed);
        let done = ok + bad;
        let mut msg = format!("{ok} ok / {bad} fail ({:.0}%)", pct(bad, done));
        if tos > 0 {
            msg.push_str(&format!("  {tos} t/o"));
        }
        suite.bar.set_message(msg);
        suite.bar.inc(1);

        let (all_ok, all_bad): (usize, usize) = self.suites.values().fold((0, 0), |(o, b), s| {
            (o + s.ok.load(Ordering::Relaxed), b + s.failed.load(Ordering::Relaxed))
        });
        self.total.set_message(format!(
            "{all_ok} ok / {all_bad} fail ({:.1}%)",
            pct(all_bad, all_ok + all_bad)
        ));
        self.total.inc(1);
    }

    /// Print above the table without tearing it.
    pub fn suspend<T>(&self, f: impl FnOnce() -> T) -> T {
        self.multi.suspend(f)
    }

    pub fn finish(&self) {
        for suite in self.suites.values() {
            suite.bar.finish_and_clear();
        }
        self.total.finish_and_clear();
    }
}

fn pct(part: usize, whole: usize) -> f64 {
    if whole == 0 { 0.0 } else { 100.0 * part as f64 / whole as f64 }
}
