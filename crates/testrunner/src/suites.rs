//! Target resolution and the live per-suite table.
//!
//! Mirrors `run_lang_tests.py`: name suites rather than paths
//! (`testrunner run go python`), and show one live row per suite instead of a
//! single undifferentiated bar, so a slow or failing language is visible while
//! the run is still going.

use crate::model::{TestExecution, TestResult};
use indicatif::{MultiProgress, ProgressBar, ProgressStyle};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};

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
    label: String,
    timing: Arc<Mutex<Timing>>,
}

#[derive(Default)]
struct Timing {
    started: Option<Instant>,
    last_finished: Option<Instant>,
    completed: usize,
    expected: usize,
}

impl Timing {
    fn start(&mut self, at: Instant) {
        self.started = Some(self.started.map_or(at, |first| first.min(at)));
    }

    fn elapsed(&self, now: Instant) -> Duration {
        let end = if self.completed == self.expected {
            self.last_finished.unwrap_or(now)
        } else { now };
        self.started.map_or(Duration::ZERO, |start| end.saturating_duration_since(start))
    }

    fn rate(&self, now: Instant) -> f64 {
        let seconds = self.elapsed(now).as_secs_f64();
        if seconds > 0.0 { self.completed as f64 / seconds } else { 0.0 }
    }
}

pub struct Summary {
    pub label: String,
    pub passed: usize,
    pub failed: usize,
    pub timeouts: usize,
    pub seconds: f64,
    pub rate: f64,
}

/// One live row per requested target, plus a total row underneath.
pub struct Table {
    multi: MultiProgress,
    suites: Vec<Suite>,
    total: ProgressBar,
    visible: bool,
}

impl Table {
    /// `visible` is decided BEFORE any bar is added. Adding first and hiding
    /// after leaves a flash of half-drawn bars on a terminal, because
    /// `MultiProgress` paints on `add`.
    pub fn new(counts: &[(String, usize)], visible: bool) -> Self {
        let multi = MultiProgress::new();
        if !visible {
            multi.set_draw_target(indicatif::ProgressDrawTarget::hidden());
        }
        let style =
            ProgressStyle::with_template(" {prefix} {bar:22.cyan/blue} {pos:>6}/{len:<6} {msg}  {suite_time}")
                .unwrap_or_else(|_| ProgressStyle::default_bar())
                .progress_chars("##-");

        let mut suites = Vec::new();
        for (lang, count) in counts {
            let timing = Arc::new(Mutex::new(Timing { expected: *count, ..Timing::default() }));
            let display_timing = timing.clone();
            let bar = multi.add(ProgressBar::new(*count as u64));
            bar.set_style(style.clone().with_key("suite_time", move |_: &indicatif::ProgressState, out: &mut dyn std::fmt::Write| {
                let timing = display_timing.lock().unwrap();
                if timing.started.is_none() {
                    let _ = out.write_str("queued");
                } else {
                    let now = Instant::now();
                    let _ = write!(out, "{:.0}/s", timing.rate(now));
                }
            }));
            // Escapes are zero-width, so the column still lines up.
            bar.set_prefix(format!("\x1b[1m{lang:<9}\x1b[0m"));
            bar.set_message("·");
            suites.push(
                Suite {
                    bar,
                    ok: AtomicUsize::new(0),
                    failed: AtomicUsize::new(0),
                    timeouts: AtomicUsize::new(0),
                    label: lang.clone(),
                    timing,
                },
            );
        }

        let grand_total: usize = counts.iter().map(|(_, count)| count).sum();
        let total = multi.add(ProgressBar::new(grand_total as u64));
        total.set_style(
            ProgressStyle::with_template(" {prefix} {bar:22.green/blue} {pos:>6}/{len:<6} {msg}")
                .unwrap_or_else(|_| ProgressStyle::default_bar())
                .progress_chars("##-"),
        );
        total.set_prefix(format!("\x1b[1m{:<9}\x1b[0m", "total"));

        Table {
            multi,
            suites,
            total,
            visible,
        }
    }

    /// Fold one finished test into its suite's row.
    pub fn start(&self, index: usize, at: Instant) {
        let Some(suite) = self.suites.get(index) else { return; };
        let first = {
            let mut timing = suite.timing.lock().unwrap();
            let first = timing.started.is_none();
            timing.start(at);
            first
        };
        if first && self.visible { suite.bar.enable_steady_tick(Duration::from_millis(250)); }
    }

    pub fn record(&self, index: usize, exec: &TestExecution, at: Instant) {
        let Some(suite) = self.suites.get(index) else {
            return;
        };
        let complete = {
            let mut timing = suite.timing.lock().unwrap();
            timing.start(at);
            timing.completed += 1;
            timing.last_finished = Some(timing.last_finished.map_or(at, |last| last.max(at)));
            timing.completed == timing.expected
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
        if complete { suite.bar.finish(); }

        let (all_ok, all_bad): (usize, usize) = self.suites.iter().fold((0, 0), |(o, b), s| {
            (
                o + s.ok.load(Ordering::Relaxed),
                b + s.failed.load(Ordering::Relaxed),
            )
        });
        self.total.set_message(format!(
            "{all_ok} ok / {all_bad} fail ({:.1}%)",
            pct(all_bad, all_ok + all_bad)
        ));
        self.total.inc(1);
    }

    pub fn summaries(&self) -> Vec<Summary> {
        self.suites.iter().filter_map(|suite| {
            let timing = suite.timing.lock().unwrap();
            timing.started?;
            let now = Instant::now();
            Some(Summary {
                label: suite.label.clone(),
                passed: suite.ok.load(Ordering::Relaxed),
                failed: suite.failed.load(Ordering::Relaxed),
                timeouts: suite.timeouts.load(Ordering::Relaxed),
                seconds: timing.elapsed(now).as_secs_f64(),
                rate: timing.rate(now),
            })
        }).collect()
    }

    /// Print above the table without tearing it.
    pub fn suspend<T>(&self, f: impl FnOnce() -> T) -> T {
        self.multi.suspend(f)
    }

    pub fn finish(&self) {
        for suite in &self.suites {
            suite.bar.finish_and_clear();
        }
        self.total.finish_and_clear();
    }
}

fn pct(part: usize, whole: usize) -> f64 {
    if whole == 0 {
        0.0
    } else {
        100.0 * part as f64 / whole as f64
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn timing_excludes_queue_time_and_freezes_at_last_completion() {
        let run_start = Instant::now();
        let mut timing = Timing { expected: 4, ..Timing::default() };
        assert_eq!(timing.elapsed(run_start + Duration::from_secs(100)), Duration::ZERO);
        timing.start(run_start + Duration::from_secs(20));
        timing.completed = 2;
        assert_eq!(timing.rate(run_start + Duration::from_secs(22)), 1.0);
        timing.completed = 4;
        timing.last_finished = Some(run_start + Duration::from_secs(24));
        assert_eq!(timing.elapsed(run_start + Duration::from_secs(100)), Duration::from_secs(4));
        assert_eq!(timing.rate(run_start + Duration::from_secs(100)), 1.0);
    }

    #[test]
    fn independently_timed_targets_and_out_of_order_starts() {
        let table = Table::new(&[("go/first".into(), 1), ("go/second".into(), 1)], false);
        let now = Instant::now();
        table.start(0, now + Duration::from_secs(3));
        table.start(0, now);
        table.start(1, now + Duration::from_secs(10));
        for (index, seconds) in [(0, 4), (1, 12)] {
            let mut timing = table.suites[index].timing.lock().unwrap();
            timing.completed = 1;
            timing.last_finished = Some(now + Duration::from_secs(seconds));
        }
        let rows = table.summaries();
        assert_eq!(rows.len(), 2);
        assert_eq!(rows[0].seconds, 4.0);
        assert_eq!(rows[0].rate, 0.25);
        assert_eq!(rows[1].seconds, 2.0);
        assert_eq!(rows[1].rate, 0.5);
        assert_eq!(rows[1].label, "go/second");
    }
}
