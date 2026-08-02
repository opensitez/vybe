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
/// When to warn that a test is still running. **Once per test, and that is the
/// whole point**: the warning says "this one is slow", which you only need to
/// be told a single time. Repeating it every 30s meant a `--timeout 300` run
/// printed ten lines about one test and buried everything else.
///
/// It is a warning, not a verdict: the test has not failed and has not timed
/// out, it is still running.
pub(crate) const SLOW_AFTER: Duration = Duration::from_secs(10);

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
    /// The deadline passed with the worker still silent — a genuine hang.
    Timeout,
    /// The worker's pipe closed, or the job could not be written to it: the
    /// process is GONE, not slow. Reporting this as a timeout was a lie with a
    /// number attached — a worker that crashed at 40s was recorded as
    /// "no result within 60s", which reads as a hang that never happened.
    Died,
}

/// Send one job and read until the result sentinel. Everything before the
/// sentinel is the program's own output, which is where the harness prints its
/// `FAIL: want [...] got [...]`.
fn dispatch(
    worker: &mut Worker,
    file: &Path,
    mode: Mode,
    timeout: Duration,
    slow: &(dyn Fn(&Path, u64) + Send + Sync),
) -> Reply {
    let tag = match mode {
        Mode::Run => "run",
        Mode::Compile => "compile",
        Mode::CompileFail => "compile-fail",
        Mode::RunFail => "run",
    };
    if writeln!(worker.stdin, "{}\t{tag}", file.display()).is_err()
        || worker.stdin.flush().is_err()
    {
        return Reply::Died;
    }

    let started = Instant::now();
    let deadline = started + timeout;
    let mut warned = false;
    let mut output = Vec::new();
    loop {
        let now = Instant::now();
        let remaining = deadline.saturating_duration_since(now);
        if remaining.is_zero() {
            return Reply::Timeout;
        }
        // Until the warning is out, wake at whichever comes first — the
        // deadline or the warning. After that there is nothing left to say, so
        // wait out the deadline in one go.
        let wait = if warned {
            remaining
        } else {
            remaining.min(
                SLOW_AFTER
                    .saturating_sub(now.duration_since(started))
                    .max(Duration::from_millis(1)),
            )
        };
        match worker.lines.recv_timeout(wait) {
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
            Err(RecvTimeoutError::Timeout) => {
                // Deadline or warning? Only the deadline ends the job.
                if Instant::now() >= deadline {
                    return Reply::Timeout;
                }
                slow(file, SLOW_AFTER.as_secs());
                warned = true;
            }
            Err(RecvTimeoutError::Disconnected) => return Reply::Died,
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
    slow: impl Fn(&Path, u64) + Send + Sync,
) -> Vec<TestExecution> {
    let queue: Arc<Mutex<VecDeque<PathBuf>>> =
        Arc::new(Mutex::new(files.iter().cloned().collect()));
    let done: Arc<Mutex<Vec<TestExecution>>> = Arc::new(Mutex::new(Vec::new()));
    let progress = Arc::new(progress);
    let slow = Arc::new(slow);

    std::thread::scope(|scope| {
        for _ in 0..jobs {
            let queue = queue.clone();
            let done = done.clone();
            let progress = progress.clone();
            let slow = slow.clone();
            scope.spawn(move || {
                let Some(mut worker) = spawn(vybex) else {
                    // Cannot start: leave the queue alone so `run_all`'s caller
                    // sees leftovers and reports a TRUNCATED run. Silently
                    // returning made a partial run look like a complete one —
                    // a `target/` clean mid-run once cut 10,385 tests to 6,281
                    // and still printed a pass rate.
                    return;
                };
                loop {
                    let Some(file) = queue.lock().unwrap().pop_front() else {
                        break;
                    };
                    let text = std::fs::read_to_string(&file).unwrap_or_default();
                    let mode = Mode::of(&text);
                    let started = Instant::now();

                    let reply = dispatch(&mut worker, &file, mode, timeout, slow.as_ref());
                    let (result, message) = match reply {
                        // In run-fail mode the program is SUPPOSED to fail.
                        Reply::Done { pass, message } if mode == Mode::RunFail => {
                            if pass {
                                (TestResult::Fail, "expected a failure, but it passed".into())
                            } else {
                                let _ = message;
                                (TestResult::Pass, String::new())
                            }
                        }
                        Reply::Done { pass: true, .. } => (TestResult::Pass, String::new()),
                        Reply::Done { message, .. } => (TestResult::Fail, message),
                        // Hung or crashed: either way the worker is replaced so
                        // the rest of the queue still drains. The two are
                        // reported differently because they are different bugs.
                        Reply::Timeout | Reply::Died => {
                            let hung = matches!(reply, Reply::Timeout);
                            let _ = worker.child.kill();
                            let _ = worker.child.wait();
                            match spawn(vybex) {
                                Some(fresh) => worker = fresh,
                                None => {
                                    queue.lock().unwrap().push_front(file);
                                    break;
                                }
                            }
                            if hung {
                                (
                                    TestResult::Timeout,
                                    format!(
                                        "no result within {}s — worker replaced",
                                        timeout.as_secs()
                                    ),
                                )
                            } else {
                                (
                                    TestResult::Error,
                                    format!(
                                        "worker died after {:.1}s without returning a result \
                                         — replaced",
                                        started.elapsed().as_secs_f64()
                                    ),
                                )
                            }
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

    // Anything still queued means workers died and could not be replaced.
    // Report it loudly: a truncated run that prints a pass rate is worse than
    // no run at all.
    let leftover = queue.lock().unwrap().len();
    if leftover > 0 {
        eprintln!(
            "[testrunner] ABORTED — {leftover} of {} test(s) never ran (no worker could start; \
             is `{}` missing or was `target/` cleaned mid-run?)",
            files.len(),
            vybex.display(),
        );
        std::process::exit(2);
    }

    Arc::try_unwrap(done).map(|m| m.into_inner().unwrap()).unwrap_or_default()
}
