//! Execute one extracted test file and turn it into a verdict.
//!
//! The verdict is the exit code and nothing else — no stdout parsing. `vybex`
//! already exits 1 on an uncaught error and 0 on a clean run, so this needs no
//! runtime change, and it is immune to the `[vybex] Project …` banner that
//! compilation prints on stdout before the program starts.
//!
//! One process per test, killed at a wall-clock deadline — the same model as
//! `testecma::runner::run_test_with_terminate`. Without the deadline a single
//! non-terminating test hangs the whole suite, and at this scale that is a
//! matter of when, not if.
//!
//! Each test carries its harness inline (spliced from `harness/<lang>/`), so a
//! case is one self-contained file — which is what lets the same path go to
//! `vybex`, to `go run`, or into the step debugger unchanged.

use crate::model::TestResult;
use std::path::Path;
use std::process::{Command, Stdio};
use std::time::{Duration, Instant};
use wait_timeout::ChildExt;

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum Mode {
    Run,
    Compile,
    CompileFail,
}

impl Mode {
    /// Read the `vybe-test-mode:` directive out of a file's header.
    pub fn of(text: &str) -> Mode {
        for line in text.lines().take(10) {
            if let Some((_, rest)) = line.split_once("vybe-test-mode:") {
                return match rest.trim() {
                    "compile" => Mode::Compile,
                    "compile-fail" => Mode::CompileFail,
                    _ => Mode::Run,
                };
            }
        }
        Mode::Run
    }
}

pub struct Outcome {
    pub result: TestResult,
    pub message: String,
    pub duration_ms: u128,
}

pub fn run_case(vybex: &Path, file: &Path, mode: Mode, timeout_secs: u64) -> Outcome {
    let mut cmd = Command::new(vybex);
    if mode != Mode::Run {
        // `-d` disassembles without running: the frontend must accept the
        // program, nothing more. That is exactly what `compile_ok` asserted.
        cmd.arg("-d");
    }
    cmd.arg(file);
    execute(cmd, mode, timeout_secs)
}

/// Run the same file under a foreign runtime — `go run`, `python3`, `node`,
/// `php`. This is the whole reason the tests are ordinary source: the
/// expectations were written against Vybe's behaviour, so a diff here is
/// evidence about the corpus, not only about us.
pub fn run_foreign(
    program: &str,
    args: &[String],
    file: &Path,
    mode: Mode,
    timeout_secs: u64,
) -> Outcome {
    let mut cmd = Command::new(program);
    cmd.args(args).arg(file);
    execute(cmd, mode, timeout_secs)
}

fn execute(mut cmd: Command, mode: Mode, timeout_secs: u64) -> Outcome {
    let started = Instant::now();
    cmd.stdout(Stdio::piped()).stderr(Stdio::piped());

    let mut child = match cmd.spawn() {
        Ok(child) => child,
        Err(e) => {
            return Outcome {
                result: TestResult::Error,
                message: format!("spawn failed: {e}"),
                duration_ms: started.elapsed().as_millis(),
            };
        }
    };

    // Drain both pipes on their own threads BEFORE waiting. A pipe holds only
    // ~64KB; `vybex -d` writes a disassembly far larger than that, so a parent
    // that waits first deadlocks against a child blocked on write — which is
    // exactly how every compile-mode test turned into a 30s timeout.
    let out_reader = child.stdout.take().map(|s| std::thread::spawn(move || drain(s)));
    let err_reader = child.stderr.take().map(|s| std::thread::spawn(move || drain(s)));

    let status = match child.wait_timeout(Duration::from_secs(timeout_secs)) {
        Ok(Some(status)) => status,
        Ok(None) => {
            let _ = child.kill();
            let _ = child.wait();
            return Outcome {
                result: TestResult::Timeout,
                message: format!("timeout after {timeout_secs}s"),
                duration_ms: started.elapsed().as_millis(),
            };
        }
        Err(e) => {
            let _ = child.kill();
            return Outcome {
                result: TestResult::Error,
                message: format!("wait failed: {e}"),
                duration_ms: started.elapsed().as_millis(),
            };
        }
    };

    let mut text = String::new();
    for reader in [out_reader, err_reader].into_iter().flatten() {
        if let Ok(chunk) = reader.join() {
            text.push_str(&chunk);
        }
    }

    let clean = status.success();
    let pass = if mode == Mode::CompileFail { !clean } else { clean };
    Outcome {
        result: if pass { TestResult::Pass } else { TestResult::Fail },
        message: if pass { String::new() } else { failure_line(&text) },
        duration_ms: started.elapsed().as_millis(),
    }
}

fn drain<R: std::io::Read>(mut stream: R) -> String {
    let mut buf = String::new();
    let _ = stream.read_to_string(&mut buf);
    buf
}

/// The first line that explains a failure. The harness prints its own
/// `FAIL: want [...] got [...]`, which is the useful one when present;
/// otherwise fall back to the runtime's own error line.
pub fn failure_line(output: &str) -> String {
    output
        .lines()
        .find(|l| l.trim_start().starts_with("FAIL: "))
        .or_else(|| {
            output
                .lines()
                .find(|l| l.contains("rror") || l.contains("panic:"))
        })
        .unwrap_or("(no output)")
        .trim()
        .chars()
        .take(200)
        .collect()
}
