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
    /// Execute it; a NON-zero exit is the pass. `must_fail(script)` asserts a
    /// wrong result really is caught — distinct from `CompileFail`, which
    /// asserts the front-end rejects the source before it ever runs.
    RunFail }

/// The exit status a test is expected to end with, from a
/// `vybe-test-exit: <n>` header. Defaults to 0.
///
/// Verdict-is-exit-code has no other way to express a test whose CORRECT
/// behaviour is a non-zero exit — a PHP script ending in `exit(1)` that real
/// php also exits 1 on. `run-fail` mode is not the same thing: it accepts ANY
/// non-zero status, so it cannot tell `exit(1)` from a crash.
pub fn expected_exit(text: &str) -> i32 {
    for line in text.lines().take(10) {
        if let Some((_, rest)) = line.split_once("vybe-test-exit:")
            && let Ok(code) = rest.trim().parse::<i32>()
        {
            return code;
        }
    }
    0
}

impl Mode {
    /// Read the `vybe-test-mode:` directive out of a file's header.
    pub fn of(text: &str) -> Mode {
        for line in text.lines().take(10) {
            if let Some((_, rest)) = line.split_once("vybe-test-mode:") {
                return match rest.trim() {
                    "compile" => Mode::Compile,
                    "compile-fail" => Mode::CompileFail,
                    "run-fail" => Mode::RunFail,
                    _ => Mode::Run };
            }
        }
        Mode::Run
    }
}

pub struct Outcome {
    pub result: TestResult,
    pub message: String,
    pub duration_ms: u128 }

pub fn run_case(
    vybex: &Path,
    file: &Path,
    mode: Mode,
    timeout_secs: u64,
    slow: &dyn Fn(u64),
) -> Outcome {
    let want_exit = std::fs::read_to_string(file).map(|t| expected_exit(&t)).unwrap_or(0);
    let mut cmd = Command::new(vybex);
    if matches!(mode, Mode::Compile | Mode::CompileFail) {
        // `-d` disassembles without running: the frontend must accept the
        // program, nothing more. That is exactly what `compile_ok` asserted.
        cmd.arg("-d");
    }
    cmd.arg(file);
    execute(cmd, mode, timeout_secs, slow, want_exit)
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
    slow: &dyn Fn(u64),
) -> Outcome {
    let want_exit = std::fs::read_to_string(file).map(|t| expected_exit(&t)).unwrap_or(0);
    let mut cmd = Command::new(program);
    cmd.args(args).arg(file);
    execute(cmd, mode, timeout_secs, slow, want_exit)
}

fn execute(
    mut cmd: Command,
    mode: Mode,
    timeout_secs: u64,
    slow: &dyn Fn(u64),
    want_exit: i32,
) -> Outcome {
    let started = Instant::now();
    cmd.stdout(Stdio::piped()).stderr(Stdio::piped());

    let mut child = match cmd.spawn() {
        Ok(child) => child,
        Err(e) => {
            return Outcome {
                result: TestResult::Error,
                message: format!("spawn failed: {e}"),
                duration_ms: started.elapsed().as_millis() };
        }
    };

    // Drain both pipes on their own threads BEFORE waiting. A pipe holds only
    // ~64KB; `vybex -d` writes a disassembly far larger than that, so a parent
    // that waits first deadlocks against a child blocked on write — which is
    // exactly how every compile-mode test turned into a 30s timeout.
    let out_reader = child.stdout.take().map(|s| std::thread::spawn(move || drain(s)));
    let err_reader = child.stderr.take().map(|s| std::thread::spawn(move || drain(s)));

    // Wait in slices rather than one blocking call, so a test that is still
    // running can say so. The warm pool does the same with `recv_timeout`; the
    // cold and foreign paths need it MORE, because `go run`/`node` hanging
    // gives no other signal at all.
    let deadline = Duration::from_secs(timeout_secs);
    let mut warned = false;
    let status = loop {
        let elapsed = started.elapsed();
        let Some(remaining) = deadline.checked_sub(elapsed) else {
            let _ = child.kill();
            let _ = child.wait();
            return Outcome {
                result: TestResult::Timeout,
                message: format!("timeout after {timeout_secs}s"),
                duration_ms: started.elapsed().as_millis() };
        };
        let wait = if warned {
            remaining
        } else {
            remaining.min(
                crate::pool::SLOW_AFTER.saturating_sub(elapsed).max(Duration::from_millis(1)),
            )
        };
        match child.wait_timeout(wait) {
            Ok(Some(status)) => break status,
            Ok(None) => {
                if started.elapsed() < deadline {
                    slow(crate::pool::SLOW_AFTER.as_secs());
                    warned = true;
                }
            }
            Err(e) => {
                let _ = child.kill();
                return Outcome {
                    result: TestResult::Error,
                    message: format!("wait failed: {e}"),
                    duration_ms: started.elapsed().as_millis() };
            }
        }
    };

    let mut text = String::new();
    for reader in [out_reader, err_reader].into_iter().flatten() {
        if let Ok(chunk) = reader.join() {
            text.push_str(&chunk);
        }
    }

    // `want_exit` is the whole point of the marker: a test that legitimately
    // ends in `exit(1)` passes on 1 and FAILS on 0, which is stricter than
    // `run-fail` (any non-zero) and than the default (must be 0).
    let clean = status.code().unwrap_or(-1) == want_exit;
    let pass = if matches!(mode, Mode::CompileFail | Mode::RunFail) {
        !status.success()
    } else {
        clean
    };
    Outcome {
        result: if pass { TestResult::Pass } else { TestResult::Fail },
        message: if pass { String::new() } else { failure_line(&text) },
        duration_ms: started.elapsed().as_millis() }
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

/// Cold path: one fresh process per test, `threads` at a time. Used for
/// `--cold` and for foreign runtimes, which have no warm protocol.
pub fn run_each(
    files: &[std::path::PathBuf],
    threads: usize,
    exec: impl Fn(&Path, Mode) -> Outcome + Sync,
    note: impl Fn(&crate::model::TestExecution) + Sync,
) -> Vec<crate::model::TestExecution> {
    use rayon::prelude::*;
    let pool = match rayon::ThreadPoolBuilder::new().num_threads(threads).build() {
        Ok(pool) => pool,
        Err(_) => return Vec::new() };
    pool.install(|| {
        files
            .par_iter()
            .map(|file| {
                let text = std::fs::read_to_string(file).unwrap_or_default();
                let mode = Mode::of(&text);
                let outcome = exec(file, mode);
                let (language, category, name) = crate::model::identify(file);
                let execution = crate::model::TestExecution {
                    path: file.clone(),
                    language,
                    category,
                    name,
                    result: outcome.result,
                    message: outcome.message,
                    duration_ms: outcome.duration_ms };
                note(&execution);
                execution
            })
            .collect()
    })
}
