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
use crate::output::{BoundedOutput, OutputEvent, read_event};
use crate::run::Mode;
use crate::schedule::Schedule;
use std::io::{BufRead, BufReader, Write};
use std::path::Path;
use std::process::{ChildStdin, Command, Stdio};
use std::sync::mpsc::{Receiver, RecvTimeoutError, SyncSender, channel, sync_channel};
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

/// Pre-booted workers kept ready so replacing a hung one is instant.
///
/// Replacing a worker used to call [`spawn`] on the consuming thread, which
/// waits for the fresh process to print `##vybe-ready`. That boot is the
/// expensive part — a one-line program costs seconds standalone, versus
/// milliseconds through a warm worker — so a hung test charged the queue the
/// timeout AND a cold start, with that thread doing nothing for both. Spares
/// move the boot off the critical path: a keeper thread keeps [`SPARES`] of
/// them warm, and a replacement is a channel receive.
///
/// A `sync_channel` is the whole mechanism: the keeper blocks on `send` while
/// the pool is full, so exactly `SPARES` idle processes exist, and it wakes to
/// brew another the moment one is taken. When `run_all` drops the receiver the
/// `send` fails and the keeper exits — no shutdown flag to get wrong.
const SPARES: usize = 2;

struct Worker {
    child: crate::safety::Child,
    stdin: ChildStdin,
    lines: Receiver<OutputEvent>,
}

fn spawn(vybex: &Path, timeout: Duration) -> Option<Worker> {
    let mut command = Command::new(vybex);
    command.arg("--worker");
    spawn_command_with_timeout(&mut command, timeout)
}

// Backpressure bounds queued allocations. Once disconnected, discard events
// but keep reading to EOF so an abandoned consumer never abandons pipe drain.
fn read_output(mut reader: impl BufRead, tx: SyncSender<OutputEvent>) {
    let mut connected = true;
    while let Ok(Some(event)) = read_event(&mut reader) {
        if connected && tx.send(event).is_err() {
            connected = false;
        }
    }
}

#[cfg(test)]
fn spawn_command(command: &mut Command) -> Option<Worker> {
    spawn_command_with_timeout(command, BOOT_TIMEOUT)
}

fn spawn_command_with_timeout(command: &mut Command, timeout: Duration) -> Option<Worker> {
    let mut child = crate::safety::spawn(
        command
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::null()),
    )
    .ok()?;
    let stdin = child.stdin.take()?;
    let stdout = child.stdout.take()?;

    let (tx, rx) = sync_channel(8);
    std::thread::spawn(move || {
        read_output(BufReader::new(stdout), tx);
    });

    let deadline = Instant::now() + timeout.min(BOOT_TIMEOUT);
    let mut ready = BoundedOutput::new();
    loop {
        match rx.recv_timeout(deadline.saturating_duration_since(Instant::now())) {
            Ok(event) => {
                ready.push(&event.bytes);
                if !event.end_of_line {
                    if Instant::now() < deadline {
                        continue;
                    }
                } else if ready.text().trim() == READY {
                    return Some(Worker {
                        child,
                        stdin,
                        lines: rx,
                    });
                }
            }
            Err(_) => {}
        }
        let _ = child.kill();
        let _ = child.wait();
        return None;
    }
}

enum Reply {
    Done {
        pass: bool,
        message: String,
        retire: bool,
    },
    /// The deadline passed with the worker still silent — a genuine hang.
    Timeout,
    /// The worker's pipe closed, or the job could not be written to it: the
    /// process is GONE, not slow. Reporting this as a timeout was a lie with a
    /// number attached — a worker that crashed at 40s was recorded as
    /// "no result within 60s", which reads as a hang that never happened.
    Died,
}

#[derive(Default)]
struct ReplyReader {
    output: BoundedOutput,
    // Only a possible sentinel prefix crosses event boundaries.
    pending: Vec<u8>,
    result: Option<BoundedOutput>,
    result_bytes: usize,
    result_line_ended: bool,
}

impl ReplyReader {
    fn event(&mut self, event: OutputEvent) -> Option<Reply> {
        if let Some(result) = &mut self.result {
            result.push(&event.bytes);
            self.result_bytes = self.result_bytes.saturating_add(event.bytes.len());
        } else {
            self.pending.extend_from_slice(&event.bytes);
            if let Some(at) = self
                .pending
                .windows(RESULT.len())
                .position(|bytes| bytes == RESULT.as_bytes())
            {
                self.output.push(&self.pending[..at]);
                let mut result = BoundedOutput::new();
                let suffix = &self.pending[at + RESULT.len()..];
                result.push(suffix);
                self.result_bytes = suffix.len();
                self.result = Some(result);
                self.pending.clear();
            } else {
                let emit = if event.end_of_line {
                    self.pending.len()
                } else {
                    self.pending.len().saturating_sub(RESULT.len() - 1)
                };
                self.output.push(&self.pending[..emit]);
                self.pending.drain(..emit);
            }
        }
        if self.result.is_some() {
            self.result_line_ended = event.end_of_line;
        }
        if event.end_of_line {
            self.finish()
        } else {
            None
        }
    }

    fn finish(&self) -> Option<Reply> {
        let result = self.result.as_ref()?.text();
        let truncated = self.result_bytes > crate::output::MAX_RETAINED_OUTPUT;
        // Match BufRead::lines: remove exactly one LF and its optional CR,
        // never bytes at the retention boundary or an unterminated EOF.
        let rest = if self.result_line_ended && !truncated {
            let line = result.strip_suffix('\n').unwrap_or(&result);
            line.strip_suffix('\r').unwrap_or(line)
        } else {
            &result
        };
        let mut parts = rest.trim_start_matches('\t').splitn(2, '\t');
        let verdict = parts.next().unwrap_or("err");
        let detail = parts.next();
        let (pass, retire) = match verdict {
            "ok" => (true, false),
            "retire-ok" => (true, true),
            "retire-err" => (false, true),
            _ => (false, false),
        };
        // A retained success token is not complete if its boundary was lost.
        let pass = pass && (!truncated || detail.is_some());
        let detail = detail.unwrap_or("");
        let message = if pass {
            String::new()
        } else {
            let printed = crate::run::failure_line(&self.output.text());
            if detail.is_empty() || printed.starts_with("FAIL: ") {
                printed
            } else {
                detail.to_string()
            }
        };
        Some(Reply::Done {
            pass,
            message,
            retire,
        })
    }
}

/// Send one job and read until the result sentinel. Everything before the
/// sentinel is the program's own output, which is where the harness prints its
/// `FAIL: want [...] got [...]`.
fn dispatch(
    worker: &mut Worker,
    file: &Path,
    text: &str,
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
    // Third field: the status the test is expected to end with
    // (`vybe-test-exit:`), so the worker can apply the same rule the cold path
    // does. Without it a test whose correct behaviour is `exit(1)` can never
    // pass under the warm pool.
    // Read from the text the CALLER already loaded. Re-reading the file here
    // meant two full reads of every test, and a job is now ~35ms — an I/O the
    // run does not need is no longer lost in the noise.
    let want_exit = crate::run::expected_exit(text);
    if writeln!(worker.stdin, "{}\t{tag}\t{want_exit}", file.display()).is_err()
        || worker.stdin.flush().is_err()
    {
        return Reply::Died;
    }

    let started = Instant::now();
    let deadline = started + timeout;
    let mut warned = false;
    let mut reader = ReplyReader::default();
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
            Ok(event) => {
                if let Some(reply) = reader.event(event) {
                    return reply;
                }
            }
            Err(RecvTimeoutError::Timeout) => {
                // Deadline or warning? Only the deadline ends the job.
                if Instant::now() >= deadline {
                    return Reply::Timeout;
                }
                slow(file, SLOW_AFTER.as_secs());
                warned = true;
            }
            Err(RecvTimeoutError::Disconnected) => return reader.finish().unwrap_or(Reply::Died),
        }
    }
}

/// Drain `schedule` through `jobs` warm workers. `progress` is called once per
/// finished test, from whichever worker thread finished it.
pub fn run_all_scheduled(
    vybex: &Path,
    schedule: Arc<Schedule>,
    jobs: usize,
    timeout: Duration,
    progress: impl Fn(&TestExecution) + Send + Sync,
    slow: impl Fn(&Path, u64) + Send + Sync,
    start: impl Fn(&Path, u32) + Send + Sync,
) -> Vec<TestExecution> {
    let total = schedule.total();
    let done: Arc<Mutex<Vec<TestExecution>>> = Arc::new(Mutex::new(Vec::new()));
    let progress = Arc::new(progress);
    let slow = Arc::new(slow);
    let start = Arc::new(start);

    // Pre-booted replacements — see [`SPARES`]. Only worth it when more than
    // one worker can hang; a single-worker run has nothing to overlap a boot
    // with.
    let (spare_tx, spare_rx) = channel::<Worker>();
    let spares: Arc<Mutex<Receiver<Worker>>> = Arc::new(Mutex::new(spare_rx));

    std::thread::scope(|scope| {
        if jobs > 1 {
            for _ in 0..SPARES {
                let spare_tx = spare_tx.clone();
                // One boot each, then the thread is done. Deliberately NOT a
                // keeper loop that refills: a thread parked on a full channel
                // never observes shutdown, and `thread::scope` would wait on it
                // forever. Booting the spares up front costs two processes and
                // needs no lifecycle at all.
                scope.spawn(move || {
                    if let Some(worker) = spawn(vybex, timeout) {
                        let _ = spare_tx.send(worker);
                    }
                });
            }
        }
        drop(spare_tx);
        let mut handles = Vec::new();
        for _ in 0..jobs {
            let spares = spares.clone();
            let schedule = schedule.clone();
            let done = done.clone();
            let progress = progress.clone();
            let slow = slow.clone();
            let start = start.clone();
            handles.push(scope.spawn(move || {
                let Some(mut worker) = spawn(vybex, timeout) else {
                    // Cannot start: leave the queue alone so `run_all`'s caller
                    // sees leftovers and reports a TRUNCATED run. Silently
                    // returning made a partial run look like a complete one —
                    // a `target/` clean mid-run once cut 10,385 tests to 6,281
                    // and still printed a pass rate.
                    return;
                };
                loop {
                    let Some(lease) = schedule.lease() else {
                        break;
                    };
                    let file = lease.path().to_owned();
                    let text = std::fs::read_to_string(&file).unwrap_or_default();
                    let mode = Mode::of(&text);
                    let started = Instant::now();

                    // A multi-unit test takes the COLD path. The warm-worker
                    // protocol sends ONE path per line, and widening it means
                    // changing `vybex --worker`; a cross-language test is rare
                    // enough that a fresh process costs nothing, and the
                    // testrunner stays the only thing that had to learn the
                    // directive.
                    if !crate::run::extra_units(&text, &file).is_empty() {
                        let outcome = crate::run::run_case(
                            vybex,
                            &file,
                            mode,
                            timeout.as_secs(),
                            &|secs| slow(&file, secs),
                            &|pid| start(&file, pid),
                        );
                        let (language, category, name) = identify(&file);
                        let exec = TestExecution {
                            path: file,
                            language,
                            category,
                            name,
                            result: outcome.result,
                            message: outcome.message,
                            duration_ms: outcome.duration_ms,
                        };
                        progress(&exec);
                        done.lock().unwrap().push(exec);
                        lease.complete();
                        continue;
                    }

                    start(&file, worker.child.id());
                    if schedule.is_cancelled() {
                        lease.retry();
                        break;
                    }
                    let reply = dispatch(&mut worker, &file, &text, mode, timeout, slow.as_ref());
                    let reply = if worker.child.memory_failure().is_some() {
                        Reply::Died
                    } else {
                        reply
                    };
                    let retire = matches!(&reply, Reply::Done { retire: true, .. });
                    let replace = retire || matches!(&reply, Reply::Timeout | Reply::Died);
                    if retire {
                        // A retiring worker has a real verdict, but must not
                        // retain guest children or receive another job. Reap
                        // its process group before reporting completion.
                        let _ = worker.child.kill();
                        let _ = worker.child.wait();
                    }
                    let (result, message) = match reply {
                        // In run-fail mode the program is SUPPOSED to fail.
                        Reply::Done { pass, message, .. } if mode == Mode::RunFail => {
                            if pass {
                                (TestResult::Fail, "expected a failure, but it passed".into())
                            } else {
                                let _ = message;
                                (TestResult::Pass, String::new())
                            }
                        }
                        Reply::Done { pass: true, .. } => (TestResult::Pass, String::new()),
                        Reply::Done { message, .. } => (TestResult::Fail, message),
                        // Reap before publishing; replacement must wait until
                        // this verdict releases its suite's admission slot.
                        Reply::Timeout | Reply::Died => {
                            let hung = matches!(reply, Reply::Timeout);
                            let memory_failure = worker.child.memory_failure();
                            let _ = worker.child.kill();
                            let _ = worker.child.wait();
                            if let Some(message) = memory_failure {
                                (TestResult::Error, message)
                            } else if hung {
                                (
                                    TestResult::Timeout,
                                    format!(
                                        "no result within {}s — worker stopped",
                                        timeout.as_secs()
                                    ),
                                )
                            } else {
                                (
                                    TestResult::Error,
                                    format!(
                                        "worker died after {:.1}s without returning a result \
                                         — worker stopped",
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
                    lease.complete();
                    if replace {
                        if schedule.is_complete() || schedule.is_cancelled() {
                            break;
                        }
                        let ready = spares.lock().unwrap().try_recv().ok();
                        match ready.or_else(|| spawn(vybex, timeout)) {
                            Some(fresh) => worker = fresh,
                            // Keep the completed verdict even if replacement
                            // fails. The unchanged queue is reported truncated.
                            None => break,
                        }
                    }
                }
                let _ = worker.child.kill();
                let _ = worker.child.wait();
            }));
        }
        // Join before draining: while a worker thread is alive it may still
        // claim a spare, and a spare killed out from under it would be handed
        // over already dead.
        for handle in handles {
            let _ = handle.join();
        }
        // Unclaimed spares are idle processes holding a booted VM. Nothing else
        // reaps them — dropping a `Child` does not kill it on Unix, so without
        // this every run would leave `SPARES` orphaned `vybex --worker`
        // processes behind.
        if let Ok(rx) = spares.lock() {
            while let Ok(mut spare) = rx.try_recv() {
                let _ = spare.child.kill();
                let _ = spare.child.wait();
            }
        }
    });

    // Let the scheduled caller drain its reporter before declaring truncation.
    // The flat wrapper retains its historical nonzero-exit behavior.
    let leftover = total.saturating_sub(done.lock().unwrap().len());
    if leftover > 0 {
        schedule.cancel();
    }

    Arc::try_unwrap(done)
        .map(|m| m.into_inner().unwrap())
        .unwrap_or_default()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::output::{MAX_RETAINED_OUTPUT, OUTPUT_CHUNK_BYTES};
    use std::io::Cursor;
    use std::path::PathBuf;
    use std::sync::atomic::{AtomicUsize, Ordering};

    /// Drain `files` as one schedule; every file must return a verdict.
    fn run_all(
        vybex: &Path,
        files: &[PathBuf],
        jobs: usize,
        timeout: Duration,
        progress: impl Fn(&TestExecution) + Send + Sync,
        slow: impl Fn(&Path, u64) + Send + Sync,
        start: impl Fn(&Path, u32) + Send + Sync,
    ) -> Vec<TestExecution> {
        let results = run_all_scheduled(
            vybex,
            Arc::new(Schedule::new(vec![files.to_vec()], 1)),
            jobs,
            timeout,
            progress,
            slow,
            start,
        );
        assert_eq!(results.len(), files.len(), "every test returns a verdict");
        results
    }

    #[test]
    fn result_sentinel_survives_every_fragment_boundary_and_binary_output() {
        let bytes = b"\xffprefix##vybe-result\terr\tdetail\r\n";
        for split in 0..bytes.len() {
            let mut reader = ReplyReader::default();
            assert!(
                reader
                    .event(OutputEvent {
                        bytes: bytes[..split].to_vec(),
                        end_of_line: false,
                    })
                    .is_none()
            );
            let reply = reader.event(OutputEvent {
                bytes: bytes[split..].to_vec(),
                end_of_line: true,
            });
            assert!(
                matches!(reply, Some(Reply::Done { pass: false, message, .. }) if message == "detail")
            );
        }
    }

    #[test]
    fn huge_prefix_and_result_detail_remain_bounded_without_losing_verdict() {
        let mut reader = ReplyReader::default();
        for _ in 0..128 {
            assert!(
                reader
                    .event(OutputEvent {
                        bytes: vec![b'x'; OUTPUT_CHUNK_BYTES],
                        end_of_line: false,
                    })
                    .is_none()
            );
            assert!(reader.pending.len() < RESULT.len());
        }
        assert!(
            reader
                .event(OutputEvent {
                    bytes: b"##vybe-result\terr\t".to_vec(),
                    end_of_line: false,
                })
                .is_none()
        );
        for _ in 0..128 {
            assert!(
                reader
                    .event(OutputEvent {
                        bytes: vec![b'd'; OUTPUT_CHUNK_BYTES],
                        end_of_line: false,
                    })
                    .is_none()
            );
        }
        assert_eq!(reader.output.text().len(), MAX_RETAINED_OUTPUT);
        assert!(
            matches!(reader.finish(), Some(Reply::Done { pass: false, message, .. })
            if !message.is_empty() && message.len() <= MAX_RETAINED_OUTPUT)
        );
    }

    #[test]
    fn printed_failure_beats_worker_detail_and_eof_can_finish_result() {
        let mut reader = ReplyReader::default();
        assert!(
            reader
                .event(OutputEvent {
                    bytes: b"FAIL: synthetic\n".to_vec(),
                    end_of_line: true,
                })
                .is_none()
        );
        assert!(
            reader
                .event(OutputEvent {
                    bytes: b"tail##vybe-result\terr\truntime detail".to_vec(),
                    end_of_line: false,
                })
                .is_none()
        );
        assert!(
            matches!(reader.finish(), Some(Reply::Done { pass: false, message, .. })
            if message == "FAIL: synthetic")
        );
    }

    #[test]
    fn malformed_or_truncated_verdict_cannot_become_ok() {
        for suffix in [b"\tok\r\r\n".as_slice(), b"\tok\r".as_slice()] {
            let mut reader = ReplyReader::default();
            let mut bytes = RESULT.as_bytes().to_vec();
            bytes.extend_from_slice(suffix);
            let reply = reader
                .event(OutputEvent {
                    bytes,
                    end_of_line: suffix.ends_with(b"\n"),
                })
                .or_else(|| reader.finish());
            assert!(matches!(reply, Some(Reply::Done { pass: false, .. })));
        }

        let mut reader = ReplyReader::default();
        reader.event(OutputEvent {
            bytes: RESULT.as_bytes().to_vec(),
            end_of_line: false,
        });
        let mut suffix = vec![b'\t'; MAX_RETAINED_OUTPUT - 2];
        suffix.extend_from_slice(b"okBAD\n");
        let mut reply = None;
        for bytes in suffix.chunks(OUTPUT_CHUNK_BYTES) {
            reply = reader.event(OutputEvent {
                bytes: bytes.to_vec(),
                end_of_line: bytes.ends_with(b"\n"),
            });
        }
        assert!(matches!(reply, Some(Reply::Done { pass: false, .. })));
    }

    #[test]
    fn retirement_verdicts_require_exact_tokens_across_fragment_boundaries() {
        for (token, pass, retire) in [
            ("ok", true, false),
            ("err", false, false),
            ("retire-ok", true, true),
            ("retire-err", false, true),
            ("retire-okBAD", false, false),
            ("retire-errBAD", false, false),
            ("retire", false, false),
            ("retire-ok ", false, false),
            ("retire-ok\r", false, false),
        ] {
            let bytes = format!("output##vybe-result\t{token}\tdetail\r\n").into_bytes();
            for split in 0..bytes.len() {
                let mut reader = ReplyReader::default();
                assert!(
                    reader
                        .event(OutputEvent {
                            bytes: bytes[..split].to_vec(),
                            end_of_line: false,
                        })
                        .is_none()
                );
                let reply = reader.event(OutputEvent {
                    bytes: bytes[split..].to_vec(),
                    end_of_line: true,
                });
                assert!(
                    matches!(reply, Some(Reply::Done {
                    pass: actual_pass, retire: actual_retire, message,
                }) if actual_pass == pass && actual_retire == retire
                    && message == if pass { "" } else { "detail" }),
                    "{token} at {split}"
                );
            }
        }
    }

    #[test]
    fn retirement_verdict_at_eof_keeps_the_real_result() {
        for (token, pass) in [("retire-ok", true), ("retire-err", false)] {
            let mut reader = ReplyReader::default();
            assert!(
                reader
                    .event(OutputEvent {
                        bytes: format!("##vybe-result\t{token}").into_bytes(),
                        end_of_line: false,
                    })
                    .is_none()
            );
            assert!(matches!(reader.finish(), Some(Reply::Done {
                pass: actual, retire: true, ..
            }) if actual == pass));
        }
    }

    #[test]
    fn retirement_verdicts_keep_bounded_detail_and_reject_truncated_tokens() {
        for (token, pass) in [("retire-ok", true), ("retire-err", false)] {
            let mut reader = ReplyReader::default();
            reader.event(OutputEvent {
                bytes: format!("##vybe-result\t{token}\t").into_bytes(),
                end_of_line: false,
            });
            for _ in 0..128 {
                assert!(
                    reader
                        .event(OutputEvent {
                            bytes: vec![b'd'; OUTPUT_CHUNK_BYTES],
                            end_of_line: false,
                        })
                        .is_none()
                );
            }
            let reply = reader.event(OutputEvent {
                bytes: b"\n".to_vec(),
                end_of_line: true,
            });
            assert!(
                matches!(reply, Some(Reply::Done { pass: actual, retire: true, message })
                if actual == pass && message.len() <= MAX_RETAINED_OUTPUT
                    && (pass || !message.is_empty()))
            );
        }
        let mut reader = ReplyReader::default();
        reader.event(OutputEvent {
            bytes: b"##vybe-result\tretire-ok".to_vec(),
            end_of_line: false,
        });
        for _ in 0..128 {
            reader.event(OutputEvent {
                bytes: vec![b'x'; OUTPUT_CHUNK_BYTES],
                end_of_line: false,
            });
        }
        assert!(matches!(
            reader.finish(),
            Some(Reply::Done {
                pass: false,
                retire: false,
                ..
            })
        ));
    }

    struct CountedReader {
        cursor: Cursor<Vec<u8>>,
        consumed: Arc<AtomicUsize>,
    }

    impl std::io::Read for CountedReader {
        fn read(&mut self, buffer: &mut [u8]) -> std::io::Result<usize> {
            std::io::Read::read(&mut self.cursor, buffer)
        }
    }

    impl BufRead for CountedReader {
        fn fill_buf(&mut self) -> std::io::Result<&[u8]> {
            self.cursor.fill_buf()
        }
        fn consume(&mut self, amount: usize) {
            self.cursor.consume(amount);
            self.consumed.fetch_add(amount, Ordering::SeqCst);
        }
    }

    #[test]
    fn bounded_queue_backpressures_then_drains_after_receiver_disconnects() {
        let consumed = Arc::new(AtomicUsize::new(0));
        let total = MAX_RETAINED_OUTPUT * 16;
        let reader = CountedReader {
            cursor: Cursor::new(vec![b'x'; total]),
            consumed: consumed.clone(),
        };
        let (tx, rx) = sync_channel(1);
        let (done_tx, done_rx) = channel();
        let handle = std::thread::spawn(move || {
            read_output(reader, tx);
            let _ = done_tx.send(());
        });
        // With no consumer, even the read-ahead fragment stays bounded.
        assert!(done_rx.recv_timeout(Duration::from_millis(20)).is_err());
        assert!(consumed.load(Ordering::SeqCst) <= OUTPUT_CHUNK_BYTES * 2);
        drop(rx);
        done_rx.recv_timeout(Duration::from_secs(5)).unwrap();
        handle.join().unwrap();
        assert_eq!(consumed.load(Ordering::SeqCst), total);
    }

    #[cfg(unix)]
    fn fake_worker(body: &str) -> Worker {
        let mut command = Command::new("/bin/sh");
        command
            .arg("-c")
            .arg(format!("printf '##vybe-ready\\n'; {body}"));
        command.env("FLOOD", "x".repeat(1024));
        spawn_command(&mut command).expect("shell fakeworker should boot")
    }

    #[cfg(unix)]
    struct FakeRuntime {
        directory: PathBuf,
        executable: PathBuf,
    }

    #[cfg(unix)]
    impl FakeRuntime {
        fn new(script: &str) -> Self {
            use std::os::unix::fs::PermissionsExt;
            static NEXT: AtomicUsize = AtomicUsize::new(0);
            let directory = std::env::temp_dir().join(format!(
                "testrunner-pool-cleanup-{}-{}-{}",
                std::process::id(),
                std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .unwrap()
                    .as_nanos(),
                NEXT.fetch_add(1, Ordering::Relaxed),
            ));
            std::fs::create_dir(&directory).unwrap();
            let executable = directory.join("fake-vybex");
            std::fs::write(&executable, script).unwrap();
            std::fs::set_permissions(&executable, std::fs::Permissions::from_mode(0o700)).unwrap();
            Self {
                directory,
                executable,
            }
        }

        fn cleanup_failure() -> Self {
            Self::new(
                r#"#!/bin/sh
[ "$1" = --worker ] || exit 90
printf '##vybe-ready\n'
while IFS= read -r job; do
    case "$job" in
        *missing_verdict.c*)
            printf 'program finished; cleanup failed\n'
            exit 23
            ;;
        *following_case.c*) printf '##vybe-result\tok\n' ;;
        *) exit 91 ;;
    esac
done
"#,
            )
        }

        fn case(&self, name: &str, text: &str) -> PathBuf {
            let path = self.directory.join(name);
            std::fs::write(&path, text).unwrap();
            path
        }
    }

    #[cfg(unix)]
    impl Drop for FakeRuntime {
        fn drop(&mut self) {
            let _ = std::fs::remove_dir_all(&self.directory);
        }
    }

    #[cfg(unix)]
    fn process_group_is_reaped(pid: u32) -> bool {
        let absent = |pid| {
            (unsafe { libc::kill(pid, 0) }) == -1
                && std::io::Error::last_os_error().raw_os_error() == Some(libc::ESRCH)
        };
        if !absent(pid as libc::pid_t) || !absent(-(pid as libc::pid_t)) {
            return false;
        }
        let mut status = 0;
        (unsafe { libc::waitpid(pid as libc::pid_t, &mut status, libc::WNOHANG) }) == -1
            && std::io::Error::last_os_error().raw_os_error() == Some(libc::ECHILD)
    }

    #[cfg(unix)]
    #[test]
    fn fakeworker_retirement_preserves_verdict_and_reaps_before_progress_and_replacement() {
        for (token, header, expected, expected_message) in [
            ("retire-ok", "", TestResult::Pass, ""),
            ("retire-err", "", TestResult::Fail, "guest failed"),
            ("retire-ok", "// vybe-test-exit: 1\n", TestResult::Pass, ""),
            (
                "retire-ok",
                "// vybe-test-mode: compile-fail\n",
                TestResult::Pass,
                "",
            ),
            (
                "retire-err",
                "// vybe-test-mode: run-fail\n",
                TestResult::Pass,
                "",
            ),
            (
                "retire-ok",
                "// vybe-test-mode: run-fail\n",
                TestResult::Fail,
                "expected a failure, but it passed",
            ),
        ] {
            let detail = if token == "retire-err" {
                "\\tguest failed"
            } else {
                ""
            };
            let runtime = FakeRuntime::new(&format!(
                r#"#!/bin/sh
[ "$1" = --worker ] || exit 90
printf '##vybe-ready\n'
while IFS= read -r job; do
    case "$job" in
        *retiring_case.c*) printf '##vybe-result\t{token}{detail}\n' ;;
        *following_case.c*) printf '##vybe-result\tok\n' ;;
        *) exit 91 ;;
    esac
done
"#
            ));
            // Deliberately stay alive after the retire marker. Even a worker
            // that fails to exit must be killed/reaped by the supervisor.
            let retired = runtime.case("retiring_case.c", header);
            let following = runtime.case("following_case.c", "");
            let starts: Mutex<Vec<(PathBuf, u32)>> = Mutex::new(Vec::new());
            let progress = Mutex::new(Vec::new());
            let results = run_all(
                &runtime.executable,
                &[retired.clone(), following.clone()],
                1,
                Duration::from_secs(2),
                |execution| {
                    let reaped = if execution.path == retired {
                        let pid = starts.lock().unwrap()[0].1;
                        process_group_is_reaped(pid)
                    } else {
                        true
                    };
                    progress.lock().unwrap().push((
                        execution.path.clone(),
                        execution.result,
                        reaped,
                    ));
                },
                |_, _| {},
                |path, pid| starts.lock().unwrap().push((path.to_owned(), pid)),
            );
            assert_eq!(results.len(), 2);
            assert_eq!(results[0].path, retired);
            assert_eq!(results[0].result, expected, "{token}: {header}");
            assert_eq!(results[0].message, expected_message);
            assert_eq!(results[1].path, following);
            assert_eq!(results[1].result, TestResult::Pass);
            assert_eq!(
                *progress.lock().unwrap(),
                vec![
                    (retired, expected, true),
                    (following, TestResult::Pass, true),
                ],
                "retirement must be reaped before reporting its result"
            );
            let starts = starts.lock().unwrap();
            assert_eq!(starts.len(), 2);
            assert_ne!(starts[0].1, starts[1].1);
            for (_, pid) in starts.iter() {
                assert!(process_group_is_reaped(*pid), "worker/group {pid} leaked");
            }
        }
    }

    #[cfg(unix)]
    #[test]
    fn fakeworker_exit_without_verdict_is_died_not_done_or_timeout() {
        // A successful OS exit is not a successful protocol verdict either.
        for status in [0, 23] {
            let mut worker = fake_worker(&format!(
                "IFS= read -r job; printf 'program finished; no verdict\\n'; exit {status}",
            ));
            let reply = dispatch(
                &mut worker,
                Path::new("synthetic"),
                "// vybe-test-exit: 1\n",
                Mode::Run,
                Duration::from_secs(2),
                &|_, _| {},
            );
            assert!(matches!(reply, Reply::Died));
            assert_eq!(worker.child.wait().unwrap().code(), Some(status));
        }
    }

    #[cfg(unix)]
    #[test]
    fn fakeworker_cleanup_barrier_precedes_verdict_and_idle_worker_can_be_reused() {
        // The second input is an explicit teardown release, not a timed sleep.
        // This tests the protocol barrier, not actual VM memory reclamation.
        let mut worker = fake_worker(
            "IFS= read -r job; printf 'program finished\\n'; \
             IFS= read -r release; [ \"$release\" = cleanup-complete ] || exit 92; \
             printf '##vybe-result\\tok\\n'; \
             while IFS= read -r job; do printf '##vybe-result\\tok\\n'; done",
        );
        let pid = worker.child.id();
        writeln!(worker.stdin, "first\trun\t0").unwrap();
        worker.stdin.flush().unwrap();
        let mut reader = ReplyReader::default();
        let output = worker.lines.recv_timeout(Duration::from_secs(2)).unwrap();
        assert!(
            reader.event(output).is_none(),
            "main returning is not a verdict"
        );
        assert!(matches!(
            worker.lines.try_recv(),
            Err(std::sync::mpsc::TryRecvError::Empty)
        ));

        writeln!(worker.stdin, "cleanup-complete").unwrap();
        worker.stdin.flush().unwrap();
        let verdict = worker.lines.recv_timeout(Duration::from_secs(2)).unwrap();
        assert!(matches!(
            reader.event(verdict),
            Some(Reply::Done { pass: true, .. })
        ));
        assert!(worker.child.try_wait().unwrap().is_none());
        assert!(matches!(
            worker.lines.try_recv(),
            Err(std::sync::mpsc::TryRecvError::Empty)
        ));
        assert!(matches!(
            dispatch(
                &mut worker,
                Path::new("following"),
                "",
                Mode::Run,
                Duration::from_secs(2),
                &|_, _| {},
            ),
            Reply::Done { pass: true, .. }
        ));
        assert_eq!(
            worker.child.id(),
            pid,
            "successful cleanup permits warm reuse"
        );
        let _ = worker.child.kill();
        let _ = worker.child.wait();
    }

    #[cfg(unix)]
    #[test]
    fn fakeworker_pool_replaces_cleanup_exit_and_completes_following_case() {
        for header in [
            "",
            "// vybe-test-exit: 1\n",
            "// vybe-test-mode: run-fail\n",
            "// vybe-test-mode: compile\n",
            "// vybe-test-mode: compile-fail\n",
        ] {
            let runtime = FakeRuntime::cleanup_failure();
            let failed = runtime.case("missing_verdict.c", header);
            let following = runtime.case("following_case.c", "");
            let starts = Mutex::new(Vec::new());
            let progress = Mutex::new(Vec::new());
            // One consumer forces the second case through the replacement
            // path, not through an unrelated worker or a speculative spare.
            let results = run_all(
                &runtime.executable,
                &[failed.clone(), following.clone()],
                1,
                Duration::from_secs(2),
                |execution| {
                    progress
                        .lock()
                        .unwrap()
                        .push((execution.path.clone(), execution.result))
                },
                |_, _| panic!("a fake cleanup exit must not produce a slow warning"),
                |path, pid| starts.lock().unwrap().push((path.to_owned(), pid)),
            );
            assert_eq!(results.len(), 2, "{header}: {results:?}");
            assert_eq!(results[0].path, failed);
            assert_eq!(
                results[0].result,
                TestResult::Error,
                "{header}: {results:?}"
            );
            assert!(results[0].message.contains("without returning a result"));
            assert!(results[0].message.contains("worker stopped"));
            assert_eq!(results[1].path, following);
            assert_eq!(results[1].result, TestResult::Pass, "{header}: {results:?}");
            assert!(results[1].message.is_empty());
            assert_eq!(
                *progress.lock().unwrap(),
                vec![
                    (failed.clone(), TestResult::Error),
                    (following.clone(), TestResult::Pass),
                ]
            );
            let starts = starts.lock().unwrap();
            assert_eq!(starts.len(), 2);
            assert_eq!(starts[0].0, failed);
            assert_eq!(starts[1].0, following);
            assert_ne!(
                starts[0].1, starts[1].1,
                "the retired worker must not get another job"
            );
        }
    }

    #[cfg(unix)]
    #[test]
    fn fakeworker_handles_multiple_huge_newline_free_jobs() {
        let mut worker = fake_worker(
            "while IFS= read -r job; do i=0; while [ \"$i\" -lt 256 ]; do \
             printf '%s' \"$FLOOD\"; i=$((i+1)); done; \
             printf '##vybe-result\\tok\\n'; done",
        );
        for _ in 0..2 {
            let reply = dispatch(
                &mut worker,
                Path::new("synthetic"),
                "",
                Mode::Run,
                Duration::from_secs(5),
                &|_, _| {},
            );
            assert!(matches!(reply, Reply::Done { pass: true, .. }));
        }
        let _ = worker.child.kill();
        let _ = worker.child.wait();
    }

    #[cfg(unix)]
    #[test]
    fn fakeworker_output_flood_still_observes_dispatch_deadline() {
        let mut worker = fake_worker("IFS= read -r job; while :; do printf '%s' \"$FLOOD\"; done");
        let started = Instant::now();
        let reply = dispatch(
            &mut worker,
            Path::new("synthetic"),
            "",
            Mode::Run,
            Duration::from_millis(100),
            &|_, _| {},
        );
        let _ = worker.child.kill();
        let _ = worker.child.wait();
        assert!(matches!(reply, Reply::Timeout));
        assert!(started.elapsed() < Duration::from_secs(5));
    }

    #[cfg(unix)]
    #[test]
    fn scheduled_fakeworker_timeout_does_not_hold_back_next_suite() {
        let runtime = FakeRuntime::new(
            r#"#!/bin/sh
[ "$1" = --worker ] || exit 90
printf '##vybe-ready\n'
while IFS= read -r job; do
    case "$job" in
        *slow_A.c*) sleep 20 ;;
        *fast_B.c*|*fast_C.c*) printf '##vybe-result\tok\n' ;;
        *) exit 91 ;;
    esac
done
"#,
        );
        let a = runtime.case("slow_A.c", "");
        let b = runtime.case("fast_B.c", "");
        let c = runtime.case("fast_C.c", "");
        let schedule = Arc::new(Schedule::new(
            vec![vec![a.clone()], vec![b], vec![c.clone()]],
            2,
        ));
        let reported = Mutex::new(Vec::new());
        let active = AtomicUsize::new(0);
        let peak = AtomicUsize::new(0);
        let started = Instant::now();
        let results = run_all_scheduled(
            &runtime.executable,
            schedule.clone(),
            2,
            Duration::from_secs(1),
            |execution| {
                active.fetch_sub(1, Ordering::SeqCst);
                reported.lock().unwrap().push(execution.path.clone());
            },
            |_, _| {},
            |_, _| {
                let count = active.fetch_add(1, Ordering::SeqCst) + 1;
                peak.fetch_max(count, Ordering::SeqCst);
            },
        );
        assert_eq!(results.len(), 3);
        assert!(schedule.is_complete());
        assert!(
            results
                .iter()
                .any(|result| result.path == a && result.result == TestResult::Timeout)
        );
        assert!(
            results
                .iter()
                .any(|result| result.path == c && result.result == TestResult::Pass)
        );
        let reported = reported.lock().unwrap();
        assert!(
            reported.iter().position(|path| path == &c).unwrap()
                < reported.iter().position(|path| path == &a).unwrap()
        );
        assert!(peak.load(Ordering::SeqCst) <= 2);
        assert_eq!(active.load(Ordering::SeqCst), 0);
        assert!(started.elapsed() < Duration::from_secs(5));
    }

    #[cfg(unix)]
    #[test]
    fn replacement_boot_failure_keeps_timeout_or_death_verdict_and_releases_lease() {
        for (body, expected) in [
            ("while :; do :; done", TestResult::Timeout),
            ("exit 23", TestResult::Error),
        ] {
            let runtime = FakeRuntime::new(&format!(
                r#"#!/bin/sh
[ "$1" = --worker ] || exit 90
directory="${{0%/*}}"
if [ -e "$directory/booted" ]; then
    : > "$directory/replacement-failed"
    exit 88
fi
: > "$directory/booted"
printf '##vybe-ready\n'
IFS= read -r job
{body}
"#
            ));
            let failed = runtime.case("failed.c", "");
            let pending = runtime.case("pending.c", "");
            let schedule = Arc::new(Schedule::new(vec![vec![failed.clone()], vec![pending]], 1));
            let starts = Mutex::new(Vec::new());
            let reported = Mutex::new(Vec::new());
            let results = run_all_scheduled(
                &runtime.executable,
                schedule.clone(),
                1,
                // This regression tests post-dispatch verdict retention, not
                // 100ms shell startup under the parallel unit-test cohort.
                Duration::from_secs(1),
                |execution| {
                    // The failure is published after reap, before any boot attempt.
                    assert!(!runtime.directory.join("replacement-failed").exists());
                    assert!(process_group_is_reaped(starts.lock().unwrap()[0]));
                    reported.lock().unwrap().push(execution.path.clone());
                },
                |_, _| {},
                |_, pid| starts.lock().unwrap().push(pid),
            );
            assert_eq!(
                starts.lock().unwrap().len(),
                1,
                "fixture never reached dispatch; initial readiness failed"
            );
            assert_eq!(results.len(), 1);
            assert_eq!(results[0].path, failed);
            assert_eq!(results[0].result, expected);
            assert!(results[0].message.contains("worker stopped"));
            assert_eq!(reported.lock().unwrap().as_slice(), &[failed.clone()]);
            assert!(runtime.directory.join("replacement-failed").exists());
            assert_eq!(starts.lock().unwrap().len(), 1);
            assert_eq!(schedule.complete(&failed), None);
            assert!(schedule.is_cancelled());
            assert!(!schedule.is_complete());
        }
    }

    #[cfg(unix)]
    #[test]
    fn readiness_and_unclaimed_spare_boots_observe_short_test_timeout() {
        let runtime = FakeRuntime::new(
            r#"#!/bin/sh
[ "$1" = --worker ] || exit 90
IFS= read -r never_sent
"#,
        );
        let pending = runtime.case("pending.c", "");
        let schedule = Arc::new(Schedule::new(vec![vec![pending]], 1));
        let started = Instant::now();
        let results = run_all_scheduled(
            &runtime.executable,
            schedule.clone(),
            2,
            Duration::from_millis(100),
            |_| panic!("unready worker returned a verdict"),
            |_, _| {},
            |_, _| panic!("unready worker received a case"),
        );
        assert!(results.is_empty());
        assert!(schedule.is_cancelled());
        assert!(!schedule.is_complete());
        assert!(started.elapsed() < Duration::from_secs(2));
    }

    #[cfg(unix)]
    #[test]
    fn twelve_workers_finish_all_fast_cases_before_one_hung_case_times_out() {
        let runtime = FakeRuntime::new(
            r#"#!/bin/sh
[ "$1" = --worker ] || exit 90
printf '##vybe-ready\n'
while IFS= read -r job; do
    case "$job" in
        *hung.c*) IFS= read -r never_sent ;;
        *fast_*) printf '##vybe-result\tok\n' ;;
        *) exit 91 ;;
    esac
done
"#,
        );
        let hung = runtime.case("hung.c", "");
        let mut files = vec![hung.clone()];
        for index in 0..39 {
            files.push(runtime.case(&format!("fast_{index:02}.c"), ""));
        }
        let schedule = Arc::new(Schedule::new(vec![files], 1));
        let gate = (Mutex::new(0usize), std::sync::Condvar::new());
        let active = AtomicUsize::new(0);
        let peak = AtomicUsize::new(0);
        let reported = Mutex::new(Vec::new());
        let initial_workers = Mutex::new(std::collections::HashSet::new());
        let results = run_all_scheduled(
            &runtime.executable,
            schedule.clone(),
            12,
            Duration::from_secs(1),
            |execution| {
                active.fetch_sub(1, Ordering::SeqCst);
                reported
                    .lock()
                    .unwrap()
                    .push((execution.path.clone(), execution.result));
            },
            |_, _| {},
            |_, pid| {
                let count = active.fetch_add(1, Ordering::SeqCst) + 1;
                peak.fetch_max(count, Ordering::SeqCst);
                let (lock, changed) = &gate;
                let mut started = lock.lock().unwrap();
                if *started < 12 {
                    initial_workers.lock().unwrap().insert(pid);
                    *started += 1;
                    if *started == 12 {
                        changed.notify_all();
                    } else {
                        let (ready, _) = changed
                            .wait_timeout_while(started, Duration::from_secs(3), |started| {
                                *started < 12
                            })
                            .unwrap();
                        started = ready;
                    }
                    // Dispatch deadlines begin after this callback returns.
                    assert_eq!(*started, 12, "not all twelve fakeworkers booted");
                }
            },
        );
        assert_eq!(results.len(), 40);
        assert!(schedule.is_complete());
        assert!(!schedule.is_cancelled());
        assert_eq!(initial_workers.lock().unwrap().len(), 12);
        assert_eq!(peak.load(Ordering::SeqCst), 12);
        assert_eq!(active.load(Ordering::SeqCst), 0);
        let reported = reported.lock().unwrap();
        assert_eq!(reported.len(), 40);
        assert!(
            reported[..39]
                .iter()
                .all(|(path, result)| path != &hung && *result == TestResult::Pass)
        );
        assert_eq!(reported[39], (hung, TestResult::Timeout));
        let unique: std::collections::HashSet<_> =
            results.iter().map(|result| &result.path).collect();
        assert_eq!(unique.len(), 40);
    }
}
