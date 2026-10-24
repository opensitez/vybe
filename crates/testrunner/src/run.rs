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
use std::path::{Path, PathBuf};
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
    RunFail,
}

/// The exit status a test is expected to end with, from a
/// `vybe-test-exit: <n>` header. Defaults to 0.
///
/// Verdict-is-exit-code has no other way to express a test whose CORRECT
/// behaviour is a non-zero exit — a PHP script ending in `exit(1)` that real
/// php also exits 1 on. `run-fail` mode is not the same thing: it accepts ANY
/// non-zero status, so it cannot tell `exit(1)` from a crash.
/// The other UNITS a test links alongside its entry file —
/// `// vybe-test-units: lib_pascal.pas other.py`.
///
/// A program is `Program { units: Vec<Bundle> }` and every unit carries its
/// OWN language, so a cross-language test is just a file list with the entry
/// LAST (`cli.rs`: "Link the other languages in first ... puts its functions
/// and classes in the shared global table"). Until this directive there was no
/// way to spell that in the corpus, which is why interop had zero tests
/// despite the whole protocol-slot design existing to serve it.
///
/// Paths resolve relative to the test file's own directory, so a pair moves
/// as a directory.
pub fn extra_units(text: &str, file: &Path) -> Vec<PathBuf> {
    let dir = file.parent().unwrap_or(Path::new("."));
    for line in text.lines().take(10) {
        if let Some((_, rest)) = line.split_once("vybe-test-units:") {
            return rest
                .split_whitespace()
                .map(|name| dir.join(name.trim_end_matches(',')))
                .collect();
        }
    }
    Vec::new()
}

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

pub fn run_case(
    vybex: &Path,
    file: &Path,
    mode: Mode,
    timeout_secs: u64,
    slow: &dyn Fn(u64),
    start: &dyn Fn(u32),
) -> Outcome {
    let text = std::fs::read_to_string(file).unwrap_or_default();
    let want_exit = expected_exit(&text);
    let mut cmd = Command::new(vybex);
    if matches!(mode, Mode::Compile | Mode::CompileFail) {
        // `-d` disassembles without running: the frontend must accept the
        // program, nothing more. That is exactly what `compile_ok` asserted.
        cmd.arg("-d");
    }
    // The ENTRY FILE IS FIRST. `single_file::load_many_program` takes
    // `entry_lang` from `paths.split_first()`, and `cli.rs` says the same —
    // "the entry file (first positional) ... the rest link alongside it".
    // Passing the library first made vybex treat PASCAL as the entry language
    // and run the PHP half as a secondary unit BEFORE it, so the Pascal
    // functions it called did not exist yet.
    cmd.arg(file);
    for unit in extra_units(&text, file) {
        cmd.arg(unit);
    }
    execute(cmd, mode, timeout_secs, slow, want_exit, start)
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
    start: &dyn Fn(u32),
) -> Outcome {
    let want_exit = std::fs::read_to_string(file)
        .map(|t| expected_exit(&t))
        .unwrap_or(0);
    let mut cmd = Command::new(program);
    cmd.args(args).arg(file);
    execute(cmd, mode, timeout_secs, slow, want_exit, start)
}

fn execute(
    mut cmd: Command,
    mode: Mode,
    timeout_secs: u64,
    slow: &dyn Fn(u64),
    want_exit: i32,
    start: &dyn Fn(u32),
) -> Outcome {
    let started = Instant::now();
    // Empty stdin, never the terminal's. A test that reads stdin must see EOF
    // and finish; inheriting fd 0 makes it block on a human who isn't there,
    // and the verdict becomes "timeout" on a machine and "passes" on a laptop
    // where someone hit return. Measured on COBOL `ACCEPT`: `< /dev/null` exits
    // 0 immediately, inherited it never returns.
    cmd.stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped());

    let mut child = match crate::safety::spawn(&mut cmd) {
        Ok(child) => child,
        Err(e) => {
            return Outcome {
                result: TestResult::Error,
                message: format!("spawn failed: {e}"),
                duration_ms: started.elapsed().as_millis(),
            };
        }
    };
    start(child.id());

    // Drain both pipes on their own threads BEFORE waiting. A pipe holds only
    // ~64KB; `vybex -d` writes a disassembly far larger than that, so a parent
    // that waits first deadlocks against a child blocked on write — which is
    // exactly how every compile-mode test turned into a 30s timeout.
    let out_reader = child
        .stdout
        .take()
        .map(|s| std::thread::spawn(move || drain(s)));
    let err_reader = child
        .stderr
        .take()
        .map(|s| std::thread::spawn(move || drain(s)));

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
                duration_ms: started.elapsed().as_millis(),
            };
        };
        let wait = if warned {
            remaining
        } else {
            remaining.min(
                crate::pool::SLOW_AFTER
                    .saturating_sub(elapsed)
                    .max(Duration::from_millis(1)),
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
                    duration_ms: started.elapsed().as_millis(),
                };
            }
        }
    };

    // Descendants may inherit output pipes. Terminate the isolated group even
    // when its leader exited, before joining pipe readers.
    let _ = child.kill();

    let mut text = String::new();
    for reader in [out_reader, err_reader].into_iter().flatten() {
        if let Ok(chunk) = reader.join() {
            text.push_str(&chunk);
        }
    }

    // `want_exit` is the whole point of the marker: a test that legitimately
    // ends in `exit(1)` passes on 1 and FAILS on 0, which is stricter than
    // `run-fail` (any non-zero) and than the default (must be 0).
    if let Some(message) = child.memory_failure() {
        return Outcome {
            result: TestResult::Error,
            message,
            duration_ms: started.elapsed().as_millis(),
        };
    }
    let clean = status.code().unwrap_or(-1) == want_exit;
    let pass = if matches!(mode, Mode::CompileFail | Mode::RunFail) {
        !status.success()
    } else {
        clean
    };
    Outcome {
        result: if pass {
            TestResult::Pass
        } else {
            TestResult::Fail
        },
        message: if pass {
            String::new()
        } else {
            failure_line(&text)
        },
        duration_ms: started.elapsed().as_millis(),
    }
}

fn drain<R: std::io::Read>(mut stream: R) -> String {
    // Keep draining after the diagnostic budget is exhausted, otherwise a
    // verbose guest blocks on its output pipe instead of reaching its verdict.
    let mut retained = Vec::new();
    let mut buffer = [0; 8192];
    loop {
        match stream.read(&mut buffer) {
            Ok(0) => break,
            Ok(n) => {
                let keep = n.min((64 * 1024usize).saturating_sub(retained.len()));
                retained.extend_from_slice(&buffer[..keep]);
            }
            Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
            Err(_) => break,
        }
    }
    String::from_utf8_lossy(&retained).into_owned()
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
pub fn run_each_scheduled(
    schedule: std::sync::Arc<crate::schedule::Schedule>,
    threads: usize,
    exec: impl Fn(&Path, Mode) -> Outcome + Sync,
    note: impl Fn(&crate::model::TestExecution) + Sync,
) -> Vec<crate::model::TestExecution> {
    use rayon::prelude::*;
    if threads == 0 {
        schedule.cancel();
        return Vec::new();
    }
    let pool = match rayon::ThreadPoolBuilder::new().num_threads(threads).build() {
        Ok(pool) => pool,
        Err(_) => {
            schedule.cancel();
            return Vec::new();
        }
    };
    pool.install(|| {
        (0..threads)
            .into_par_iter()
            .map(|_| {
                let mut results = Vec::new();
                while let Some(lease) = schedule.lease() {
                    let file = lease.path();
                    let text = std::fs::read_to_string(file).unwrap_or_default();
                    let mode = Mode::of(&text);
                    let outcome = exec(file, mode);
                    let (language, category, name) = crate::model::identify(file);
                    let execution = crate::model::TestExecution {
                        path: file.to_owned(),
                        language,
                        category,
                        name,
                        result: outcome.result,
                        message: outcome.message,
                        duration_ms: outcome.duration_ms,
                    };
                    note(&execution);
                    results.push(execution);
                    lease.complete();
                }
                results
            })
            .collect::<Vec<_>>()
            .into_iter()
            .flatten()
            .collect()
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn scheduled_cold_workers_roll_suites_while_first_case_is_blocked() {
        use crate::schedule::Schedule;
        use std::sync::{Arc, Mutex, mpsc};
        let schedule = Arc::new(Schedule::new(
            vec![
                vec![PathBuf::from("A")],
                vec![PathBuf::from("B")],
                vec![PathBuf::from("C")],
            ],
            2,
        ));
        let (started_tx, started_rx) = mpsc::channel();
        let (release_tx, release_rx) = mpsc::channel();
        let started_rx = Mutex::new(started_rx);
        let release_rx = Mutex::new(release_rx);
        let reported = Mutex::new(Vec::new());
        let copy = schedule.clone();
        let results = run_each_scheduled(
            schedule.clone(),
            2,
            |path, _| {
                match path.to_str().unwrap() {
                    "A" => {
                        started_tx.send(()).unwrap();
                        release_rx
                            .lock()
                            .unwrap()
                            .recv_timeout(Duration::from_secs(2))
                            .unwrap();
                    }
                    "B" => started_rx
                        .lock()
                        .unwrap()
                        .recv_timeout(Duration::from_secs(2))
                        .unwrap(),
                    "C" => {
                        assert_eq!(reported.lock().unwrap().as_slice(), &[PathBuf::from("B")]);
                        release_tx.send(()).unwrap();
                    }
                    _ => unreachable!(),
                }
                Outcome {
                    result: TestResult::Pass,
                    message: String::new(),
                    duration_ms: 0,
                }
            },
            |execution| {
                // This callback must precede completion's admission of C.
                if execution.path == Path::new("B") {
                    assert!(!copy.is_complete());
                }
                reported.lock().unwrap().push(execution.path.clone());
            },
        );
        assert_eq!(results.len(), 3);
        assert!(schedule.is_complete());
    }

    #[test]
    fn scheduled_cold_worker_panic_cancels_waiters_and_is_not_a_pass() {
        use std::sync::Arc;
        let schedule = Arc::new(crate::schedule::Schedule::new(
            vec![vec![PathBuf::from("panic")], vec![PathBuf::from("never")]],
            1,
        ));
        let copy = schedule.clone();
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            run_each_scheduled(copy, 2, |_, _| panic!("synthetic executor panic"), |_| {});
        }));
        assert!(result.is_err());
        assert!(schedule.is_cancelled());
        assert!(!schedule.is_complete());
    }

    #[test]
    fn output_capture_is_bounded_but_drains_to_eof() {
        let bytes = vec![b'x'; 2 * 1024 * 1024];
        let mut input = std::io::Cursor::new(bytes);
        let text = drain(&mut input);
        assert_eq!(text.len(), 64 * 1024);
        assert_eq!(input.position(), 2 * 1024 * 1024);
    }

    #[cfg(unix)]
    #[test]
    fn descendant_output_pipe_cannot_hold_up_a_completed_test() {
        let mut command = Command::new("/bin/sh");
        command.args(["-c", "sleep 20 & exit 0"]);
        let started = Instant::now();
        let result = execute(command, Mode::Run, 2, &|_| {}, 0, &|_| {});
        assert!(
            matches!(result.result, TestResult::Pass),
            "{}",
            result.message
        );
        assert!(started.elapsed() < Duration::from_secs(2));
    }

    #[cfg(unix)]
    #[test]
    fn cold_timeout_kills_the_group_and_returns() {
        let mut command = Command::new("/bin/sh");
        command.args(["-c", "sleep 20 & wait"]);
        let started = Instant::now();
        let result = execute(command, Mode::Run, 1, &|_| {}, 0, &|_| {});
        assert!(matches!(result.result, TestResult::Timeout));
        assert!(started.elapsed() < Duration::from_secs(3));
    }
}
