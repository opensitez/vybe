//! Output state owned exclusively by the reporting thread.
use std::collections::HashMap;
use std::io::Write;
use std::path::{Path, PathBuf};
use std::sync::mpsc::SyncSender;
use std::time::{Duration, Instant};

use anyhow::{Context, Result};

use crate::{model, report, style, suites};

pub enum Event {
    Start {
        path: PathBuf,
        pid: u32,
        at: Instant,
        durable: Option<SyncSender<Result<(), String>>>,
    },
    Finished { execution: model::TestExecution, at: Instant },
    Slow { path: PathBuf, seconds: u64 },
    FinishRun,
}

pub struct Config<'a> {
    pub roots: &'a [PathBuf],
    pub owner: &'a HashMap<PathBuf, usize>,
    pub results_dir: &'a Path,
    pub under: &'a str,
    pub provenance: &'a str,
    pub started: Instant,
    pub table: &'a suites::Table,
    pub progress: bool,
    pub verbose: bool,
    pub save: bool,
}

struct SuiteLog {
    path: PathBuf,
    expected: usize,
    finished: usize,
    started: Option<Instant>,
    last_finished: Option<Instant>,
    file: Option<std::fs::File>,
    report: Option<model::TestReport>,
}

pub struct State<'a> {
    config: Config<'a>,
    suites: Vec<SuiteLog>,
    journal: Option<(std::fs::File, Instant)>,
}

impl<'a> State<'a> {
    pub fn new(config: Config<'a>) -> Result<Self> {
        let mut counts = vec![0; config.roots.len()];
        for &suite in config.owner.values() { counts[suite] += 1; }
        let suites = config.roots.iter().enumerate().map(|(i, root)| SuiteLog {
            path: crate::saved_log_path(config.results_dir, root),
            expected: counts[i], finished: 0, started: None, last_finished: None, file: None,
            report: config.save.then(|| model::TestReport::new(config.under.to_owned())),
        }).collect();
        let journal = if config.save {
            std::fs::create_dir_all(config.results_dir)?;
            let stamp = std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?.as_nanos();
            let path = config.results_dir.join(format!("active-{}-{stamp}.jsonl", std::process::id()));
            let file = std::fs::OpenOptions::new().create_new(true).write(true)
                .open(&path).with_context(|| format!("creating {}", path.display()))?;
            eprintln!("[testrunner] active-test journal: {}", path.display());
            Some((file, Instant::now()))
        } else { None };
        Ok(Self { config, suites, journal })
    }

    fn open_suite(&mut self, path: &Path, at: Instant) -> Result<Option<usize>> {
        let Some(&index) = self.config.owner.get(path) else { return Ok(None); };
        self.config.table.start(index, at);
        let suite = &mut self.suites[index];
        if suite.started.is_none() {
            suite.started = Some(at);
            if self.config.save {
                if let Some(parent) = suite.path.parent() { std::fs::create_dir_all(parent)?; }
                let mut file = std::fs::File::create(&suite.path)
                    .with_context(|| format!("creating {}", suite.path.display()))?;
                writeln!(file, "running {} tests", suite.expected)?;
                writeln!(file, "{}", self.config.provenance)?;
                file.flush()?;
                std::fs::write(crate::run_marker(&suite.path), format!("{}\n", std::process::id()))?;
                suite.file = Some(file);
            }
        } else if let Some(started) = &mut suite.started {
            *started = (*started).min(at);
        }
        Ok(Some(index))
    }

    fn record(&mut self, event: &str, path: &Path, pid: Option<u32>,
        message: Option<&str>, at: Instant, durable: bool) -> Result<()> {
        if let Some((file, last_sync)) = &mut self.journal {
            let record = serde_json::json!({"event":event,"path":path,"pid":pid,
                "elapsed_ms":at.saturating_duration_since(self.config.started).as_millis(),"message":message});
            let mut bytes = serde_json::to_vec(&record)?;
            bytes.push(b'\n');
            file.write_all(&bytes)?;
            if durable || last_sync.elapsed() >= Duration::from_millis(250) {
                file.sync_data()?;
                *last_sync = Instant::now();
            }
        }
        Ok(())
    }

    pub fn handle(&mut self, event: Event) -> Result<(), String> {
        self.handle_inner(event).map_err(|error| format!("report writer: {error:#}"))
    }

    fn handle_inner(&mut self, event: Event) -> Result<()> {
        match event {
            Event::Start { path, pid, at, durable } => {
                self.open_suite(&path, at)?;
                self.record("start", &path, Some(pid), None, at, durable.is_some())?;
                if let Some(ack) = durable { let _ = ack.send(Ok(())); }
            }
            Event::Finished { execution, at } => {
                let index = self.open_suite(&execution.path, at)?;
                self.record("finish", &execution.path, None, Some(&execution.message), at, false)?;
                if let Some(index) = index { self.config.table.record(index, &execution, at); }
                if !self.config.progress {
                    println!("{}", report::cargo_verdict_line(&execution, true));
                }
                if self.config.verbose && execution.result != model::TestResult::Pass {
                    self.config.table.suspend(|| {
                        println!("FAIL {}", execution.slug());
                        println!("     {}", execution.message);
                        println!("     {}", execution.path.display());
                    });
                }
                if let Some(index) = index {
                    let suite = &mut self.suites[index];
                    suite.finished += 1;
                    suite.last_finished = Some(suite.last_finished.map_or(at, |last| last.max(at)));
                    if let Some(file) = &mut suite.file {
                        writeln!(file, "{}", report::cargo_verdict_line(&execution, false))?;
                    }
                    if let Some(report) = &mut suite.report { report.add_execution(execution); }
                    if suite.finished == suite.expected {
                        if let (Some(file), Some(report)) = (&mut suite.file, &suite.report) {
                            let elapsed = suite.last_finished.unwrap()
                                .saturating_duration_since(suite.started.unwrap()).as_secs_f64();
                            let rate = if elapsed > 0.0 { suite.finished as f64 / elapsed } else { 0.0 };
                            writeln!(file, "suite throughput: {rate:.0}/s ({} completed in {elapsed:.2}s)", suite.finished)?;
                            for line in report::cargo_tail(report, elapsed, false) { writeln!(file, "{line}")?; }
                            file.flush()?;
                            suite.file = None;
                            suite.report = None;
                            let _ = std::fs::remove_file(crate::run_marker(&suite.path));
                            eprintln!("[testrunner] saved: {}", suite.path.display());
                        }
                    }
                }
            }
            Event::Slow { path, seconds } => {
                let (language, category, name) = model::identify(&path);
                self.config.table.suspend(|| println!("{} {language}/{category}/{name} still running after {seconds}s",
                    style::yellow("warning:")));
            }
            Event::FinishRun => {
                if let Some((file, _)) = &mut self.journal { file.sync_data()?; }
                for suite in &mut self.suites {
                    if let Some(file) = &mut suite.file { file.flush()?; }
                }
            }
        }
        Ok(())
    }
}
