//! Warm worker pool — testecma's `run_reset_pool`, generalised.
//!
//! N long-lived `vybex --worker` processes pull from one shared queue. Each
//! boots a VM once and resets between jobs, so the ~90% of a run that is VM
//! setup is paid `N` times instead of once per test.
//!
//! A worker that hangs or dies is killed, its job is recorded as a timeout, and
//! a replacement is spawned — the queue keeps draining. That recovery is the
//! reason a pool is usable at all: at 6,000+ tests, one non-terminating program
//! is a certainty, not a risk.

use crate::model::{TestExecution, TestResult, identify};
use crate::run::Mode;
use std::collections::VecDeque;
use std::io::{BufRead, BufReader, Write};
use std::path::{Path, PathBuf};
use std::process::{Child, ChildStdin, Command, Stdio};
use std::sync::mpsc::{Receiver, RecvTimeoutError, channel};
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};

const READY: &str = "##vybe-ready";
const RESULT: &str = "##vybe-result";
/// A worker that hasn't finished booting in this long is broken, not slow.
const BOOT_TIMEOUT: Duration = Duration::from_secs(60);

struct Worker {
    child: Child,
    stdin: ChildStdin,
    lines: Receiver<String>,
}

fn spawn(vybex: &Path) -> Option<Worker> {
    let mut child = Command::new(vybex)
        .arg("--worker")
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .spawn()
        .ok()?;
    let stdin = child.stdin.take()?;
    let stdout = child.stdout.take()?;

    // A reader thread keeps the pipe drained; the pool never blocks the worker
    // on a full buffer, which is how the single-process path deadlocked.
    let (tx, rx) = channel();
    std::thread::spawn(move || {
        for line in BufReader::new(stdout).lines().map_while(Result::ok) {
            if tx.send(line).is_err() {
                break;
            }
        }
    });

    match rx.recv_timeout(BOOT_TIMEOUT) {
        Ok(line) if line.trim() == READY => Some(Worker { child, stdin, lines: rx }),
        _ => {
            let _ = child.kill();
            let _ = child.wait();
            None
        }
    }
}

enum Reply {
    Done { pass: bool, message: String },
    /// The worker stopped answering — it must be replaced.
    Lost,
}

/// Send one job and read until the result sentinel. Everything before the
/// sentinel is the program's own output, which is where the harness prints its
/// `FAIL: want [...] got [...]`.
fn dispatch(worker: &mut Worker, file: &Path, mode: Mode, timeout: Duration) -> Reply {
    let tag = match mode {
        Mode::Run => "run",
        Mode::Compile => "compile",
        Mode::CompileFail => "compile-fail",
    };
    if writeln!(worker.stdin, "{}\t{tag}", file.display()).is_err()
        || worker.stdin.flush().is_err()
    {
        return Reply::Lost;
    }

    let deadline = Instant::now() + timeout;
    let mut output = Vec::new();
    loop {
        let remaining = deadline.saturating_duration_since(Instant::now());
        if remaining.is_zero() {
            return Reply::Lost;
        }
        match worker.lines.recv_timeout(remaining) {
            Ok(line) => {
                if let Some(rest) = line.strip_prefix(RESULT) {
                    let mut parts = rest.trim_start_matches('\t').splitn(2, '\t');
                    let verdict = parts.next().unwrap_or("err");
                    let detail = parts.next().unwrap_or("").to_string();
                    let pass = verdict == "ok";
                    let message = if pass {
                        String::new()
                    } else if detail.is_empty() {
                        crate::run::failure_line(&output.join("\n"))
                    } else {
                        // The program's own diagnostic beats the runtime's when
                        // it printed one.
                        let printed = crate::run::failure_line(&output.join("\n"));
                        if printed.starts_with("FAIL: ") { printed } else { detail }
                    };
                    return Reply::Done { pass, message };
                }
                output.push(line);
            }
            Err(RecvTimeoutError::Timeout) | Err(RecvTimeoutError::Disconnected) => {
                return Reply::Lost;
            }
        }
    }
}

/// Drain `files` through `jobs` warm workers. `progress` is called once per
/// finished test, from whichever worker thread finished it.
pub fn run_all(
    vybex: &Path,
    files: &[PathBuf],
    jobs: usize,
    timeout: Duration,
    progress: impl Fn(&TestExecution) + Send + Sync,
) -> Vec<TestExecution> {
    let queue: Arc<Mutex<VecDeque<PathBuf>>> =
        Arc::new(Mutex::new(files.iter().cloned().collect()));
    let done: Arc<Mutex<Vec<TestExecution>>> = Arc::new(Mutex::new(Vec::new()));
    let progress = Arc::new(progress);

    std::thread::scope(|scope| {
        for _ in 0..jobs {
            let queue = queue.clone();
            let done = done.clone();
            let progress = progress.clone();
            scope.spawn(move || {
                let Some(mut worker) = spawn(vybex) else {
                    return;
                };
                loop {
                    let Some(file) = queue.lock().unwrap().pop_front() else {
                        break;
                    };
                    let text = std::fs::read_to_string(&file).unwrap_or_default();
                    let mode = Mode::of(&text);
                    let started = Instant::now();

                    let reply = dispatch(&mut worker, &file, mode, timeout);
                    let (result, message) = match reply {
                        Reply::Done { pass: true, .. } => (TestResult::Pass, String::new()),
                        Reply::Done { message, .. } => (TestResult::Fail, message),
                        Reply::Lost => {
                            // Hung or crashed: replace the worker so the rest of
                            // the queue still drains.
                            let _ = worker.child.kill();
                            let _ = worker.child.wait();
                            match spawn(vybex) {
                                Some(fresh) => worker = fresh,
                                None => {
                                    queue.lock().unwrap().push_front(file);
                                    break;
                                }
                            }
                            (
                                TestResult::Timeout,
                                format!("no result within {}s — worker replaced", timeout.as_secs()),
                            )
                        }
                    };

                    let (language, category, name) = identify(&file);
                    let exec = TestExecution {
                        path: file,
                        language,
                        category,
                        name,
                        result,
                        message,
                        duration_ms: started.elapsed().as_millis(),
                    };
                    progress(&exec);
                    done.lock().unwrap().push(exec);
                }
                let _ = worker.child.kill();
                let _ = worker.child.wait();
            });
        }
    });

    Arc::try_unwrap(done).map(|m| m.into_inner().unwrap()).unwrap_or_default()
}
