//! Bounded, nonblocking handoff to one scoped reporting thread.
//!
//! Enqueueing is not a durable commit. Events may carry their own acknowledgment
//! sender for callers explicitly opting into synchronous durability. The handler
//! owns all I/O, flushing and acknowledgment policy.

use std::fmt;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::mpsc::{SyncSender, TrySendError, sync_channel};
use std::sync::{Arc, Mutex};
use std::thread::{Scope, ScopedJoinHandle};

pub const DEFAULT_CAPACITY: usize = 256;
pub const MAX_CAPACITY: usize = 1024;

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ReporterError {
    InvalidCapacity(usize),
    Spawn(String),
    QueueFull,
    Disconnected,
    Handler(String),
    Panicked(String),
}

impl fmt::Display for ReporterError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidCapacity(capacity) => write!(
                f,
                "reporter capacity {capacity} is outside 1..={MAX_CAPACITY}"
            ),
            Self::Spawn(error) => write!(f, "cannot start reporter: {error}"),
            Self::QueueFull => write!(f, "reporter queue is full; reporting cancelled"),
            Self::Disconnected => write!(f, "reporter disconnected"),
            Self::Handler(error) => write!(f, "reporter failed: {error}"),
            Self::Panicked(error) => write!(f, "reporter panicked: {error}"),
        }
    }
}

impl std::error::Error for ReporterError {}

#[derive(Default)]
struct Failure {
    failed: AtomicBool,
    error: Mutex<Option<ReporterError>>,
}

impl Failure {
    fn record(&self, error: ReporterError) -> ReporterError {
        let mut first = self.error.lock().unwrap();
        let error = first.get_or_insert(error).clone();
        self.failed.store(true, Ordering::Release);
        error
    }

    fn get(&self) -> Option<ReporterError> {
        if self.failed.load(Ordering::Acquire) {
            self.error.lock().unwrap().clone()
        } else {
            None
        }
    }
}

#[must_use = "call finish to join the reporter and propagate reporting failures"]
pub struct Reporter<'scope, E> {
    // Keep the sender private: an escaped clone could keep finish() waiting
    // for channel closure after the producer work has already finished.
    sender: Option<SyncSender<E>>,
    failure: Arc<Failure>,
    thread: Option<ScopedJoinHandle<'scope, ()>>,
}

impl<'scope, E: Send + 'scope> Reporter<'scope, E> {
    pub fn spawn<'env, F>(
        scope: &'scope Scope<'scope, 'env>,
        capacity: usize,
        mut handler: F,
    ) -> Result<Self, ReporterError>
    where
        F: FnMut(E) -> Result<(), String> + Send + 'scope,
    {
        if !(1..=MAX_CAPACITY).contains(&capacity) {
            return Err(ReporterError::InvalidCapacity(capacity));
        }
        let (sender, receiver) = sync_channel(capacity);
        let failure = Arc::new(Failure::default());
        let writer_failure = failure.clone();
        let thread = std::thread::Builder::new()
            .name("testrunner-reporter".into())
            .spawn_scoped(scope, move || {
                for event in receiver {
                    if writer_failure.failed.load(Ordering::Acquire) {
                        // Drop queued events (including optional ack senders)
                        // while still draining so closure never requires space.
                        continue;
                    }
                    match catch_unwind(AssertUnwindSafe(|| handler(event))) {
                        Ok(Ok(())) => {}
                        Ok(Err(error)) => {
                            writer_failure.record(ReporterError::Handler(error));
                        }
                        Err(payload) => {
                            let message = if let Some(message) = payload.downcast_ref::<String>() {
                                message.clone()
                            } else if let Some(message) = payload.downcast_ref::<&str>() {
                                (*message).to_owned()
                            } else {
                                "non-string panic payload".to_owned()
                            };
                            writer_failure.record(ReporterError::Panicked(message));
                            // A panicked handler cannot be reused. Disconnect
                            // producers and drop all remaining events/acks.
                            break;
                        }
                    }
                }
            })
            .map_err(|error| ReporterError::Spawn(error.to_string()))?;
        Ok(Self {
            sender: Some(sender),
            failure,
            thread: Some(thread),
        })
    }

    /// Accept an event without waiting for queue space or any writer I/O.
    /// On error the caller should cancel scheduling; finish also returns the
    /// latched failure, so a dropped event cannot yield a successful report.
    pub fn send(&self, event: E) -> Result<(), ReporterError> {
        if let Some(error) = self.failure.get() {
            return Err(error);
        }
        match self.sender.as_ref().unwrap().try_send(event) {
            Ok(()) => Ok(()),
            Err(TrySendError::Full(_)) => Err(self.failure.record(ReporterError::QueueFull)),
            Err(TrySendError::Disconnected(_)) => {
                Err(self.failure.record(ReporterError::Disconnected))
            }
        }
    }

    pub fn failed(&self) -> bool {
        self.failure.failed.load(Ordering::Acquire)
    }

    /// After all producers have joined, enqueue a final flush/finalization
    /// event, then close and join. Only this end-of-run path waits for queue
    /// space: a temporarily full queue must not cancel normal shutdown.
    pub fn finish_with(self, event: E) -> Result<(), ReporterError> {
        if self.failure.get().is_none() {
            if self.sender.as_ref().unwrap().send(event).is_err() {
                self.failure.record(ReporterError::Disconnected);
            }
        }
        // Handler failure may race the blocking send. finish returns the
        // first latched error even if the failed writer drained the event.
        self.finish()
    }

    /// Close the queue, drain it (discarding after failure), and join the writer.
    /// The handler must terminate: this does not interrupt in-progress I/O.
    pub fn finish(mut self) -> Result<(), ReporterError> {
        drop(self.sender.take());
        if self.thread.take().unwrap().join().is_err() {
            self.failure.record(ReporterError::Panicked(
                "writer thread failed outside handler".into(),
            ));
        }
        match self.failure.get() {
            Some(error) => Err(error),
            None => Ok(()),
        }
    }
}

impl<E> Drop for Reporter<'_, E> {
    fn drop(&mut self) {
        // Scope teardown also joins if a caller returns early without finish.
        drop(self.sender.take());
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::mpsc::channel;
    use std::time::Duration;

    const DEADLINE: Duration = Duration::from_secs(5);

    #[test]
    fn reporter_can_be_shared_by_worker_callbacks() {
        fn send_sync<T: Send + Sync>() {}
        send_sync::<Reporter<'static, ()>>();
    }

    #[test]
    fn borrowed_handler_preserves_order_and_finish_drains() {
        let mut written = Vec::new();
        std::thread::scope(|scope| {
            let reporter = Reporter::spawn(scope, 8, |event| {
                written.push(event);
                Ok(())
            })
            .unwrap();
            for event in 0..8 {
                reporter.send(event).unwrap();
            }
            assert!(!reporter.failed());
            reporter.finish().unwrap();
        });
        assert_eq!(written, (0..8).collect::<Vec<_>>());
    }

    #[test]
    fn finish_with_waits_for_full_queue_and_delivers_final_event_last() {
        let mut written = Vec::new();
        std::thread::scope(|scope| {
            let (entered_tx, entered_rx) = channel();
            let (release_tx, release_rx) = channel();
            let (finishing_tx, finishing_rx) = channel();
            let written = &mut written;
            let reporter = Reporter::spawn(scope, 1, move |event| {
                if event == 0 {
                    entered_tx.send(()).map_err(|error| error.to_string())?;
                    release_rx
                        .recv_timeout(DEADLINE)
                        .map_err(|error| error.to_string())?;
                }
                written.push(event);
                Ok(())
            })
            .unwrap();
            reporter.send(0).unwrap();
            entered_rx.recv_timeout(DEADLINE).unwrap();
            reporter.send(1).unwrap();
            let finalizer = scope.spawn(move || {
                finishing_tx.send(()).unwrap();
                reporter.finish_with(2)
            });
            let finishing = finishing_rx.recv_timeout(DEADLINE);
            let _ = release_tx.send(());
            finishing.unwrap();
            assert_eq!(finalizer.join().unwrap(), Ok(()));
        });
        assert_eq!(written, vec![0, 1, 2]);
    }

    #[test]
    fn finish_with_propagates_handler_failure_while_queue_is_full() {
        let mut calls = 0;
        std::thread::scope(|scope| {
            let (entered_tx, entered_rx) = channel();
            let (release_tx, release_rx) = channel();
            let (finishing_tx, finishing_rx) = channel();
            let calls = &mut calls;
            let reporter = Reporter::spawn(scope, 1, move |_: usize| {
                *calls += 1;
                entered_tx.send(()).map_err(|error| error.to_string())?;
                release_rx
                    .recv_timeout(DEADLINE)
                    .map_err(|error| error.to_string())?;
                Err("disk failed".into())
            })
            .unwrap();
            reporter.send(0).unwrap();
            entered_rx.recv_timeout(DEADLINE).unwrap();
            reporter.send(1).unwrap();
            let finalizer = scope.spawn(move || {
                finishing_tx.send(()).unwrap();
                reporter.finish_with(2)
            });
            let finishing = finishing_rx.recv_timeout(DEADLINE);
            let _ = release_tx.send(());
            finishing.unwrap();
            assert_eq!(
                finalizer.join().unwrap(),
                Err(ReporterError::Handler("disk failed".into()))
            );
        });
        assert_eq!(
            calls, 1,
            "no finalization callback may follow handler failure"
        );
    }

    #[test]
    fn finish_with_reports_final_event_failure() {
        std::thread::scope(|scope| {
            let reporter = Reporter::spawn(scope, 1, |_: ()| Err("flush failed".into())).unwrap();
            assert_eq!(
                reporter.finish_with(()),
                Err(ReporterError::Handler("flush failed".into()))
            );
        });
    }

    #[test]
    fn slow_handler_full_queue_returns_error_before_writer_release() {
        std::thread::scope(|scope| {
            let (entered_tx, entered_rx) = channel();
            let (release_tx, release_rx) = channel();
            let (returned_tx, returned_rx) = channel();
            let reporter = Reporter::spawn(scope, 1, move |_| {
                entered_tx.send(()).map_err(|error| error.to_string())?;
                release_rx
                    .recv_timeout(DEADLINE)
                    .map_err(|error| error.to_string())?;
                Ok(())
            })
            .unwrap();
            reporter.send(0).unwrap();
            entered_rx.recv_timeout(DEADLINE).unwrap();
            reporter.send(1).unwrap();
            let producer = scope.spawn(move || {
                let result = reporter.send(2);
                returned_tx
                    .send((result.clone(), reporter.failed()))
                    .unwrap();
                reporter
            });
            // A blocking send could only complete after writer release. Always
            // release before asserting, so a regression cannot strand a scope.
            let returned = returned_rx.recv_timeout(DEADLINE);
            let _ = release_tx.send(());
            let reporter = producer.join().unwrap();
            assert_eq!(returned.unwrap(), (Err(ReporterError::QueueFull), true));
            assert_eq!(reporter.finish(), Err(ReporterError::QueueFull));
        });
    }

    #[test]
    fn handler_failure_is_latched_and_queued_acknowledgments_disconnect() {
        let mut calls = 0;
        std::thread::scope(|scope| {
            let (entered_tx, entered_rx) = channel();
            let (release_tx, release_rx) = channel();
            let (ack_tx, ack_rx) = channel::<Result<(), String>>();
            let calls = &mut calls;
            let reporter = Reporter::spawn(
                scope,
                2,
                move |_: Option<std::sync::mpsc::Sender<Result<(), String>>>| {
                    *calls += 1;
                    entered_tx.send(()).map_err(|error| error.to_string())?;
                    release_rx
                        .recv_timeout(DEADLINE)
                        .map_err(|error| error.to_string())?;
                    Err("disk failed".into())
                },
            )
            .unwrap();
            reporter.send(None).unwrap();
            entered_rx.recv_timeout(DEADLINE).unwrap();
            reporter.send(Some(ack_tx)).unwrap();
            release_tx.send(()).unwrap();
            assert_eq!(
                reporter.finish(),
                Err(ReporterError::Handler("disk failed".into()))
            );
            assert!(matches!(
                ack_rx.try_recv(),
                Err(std::sync::mpsc::TryRecvError::Disconnected)
            ));
        });
        assert_eq!(
            calls, 1,
            "failed handlers must not receive subsequent events"
        );
    }

    #[test]
    fn handler_panic_is_a_finish_error_not_a_scope_panic() {
        let result = std::thread::scope(|scope| {
            let reporter = Reporter::spawn(scope, 1, |_: ()| -> Result<(), String> {
                panic!("writer panic")
            })
            .unwrap();
            // Test-only clone: even an outstanding sender cannot keep a
            // panicked receiver alive or make finish wait for more events.
            let sender = reporter.sender.as_ref().unwrap().clone();
            reporter.send(()).unwrap();
            let result = reporter.finish();
            assert!(matches!(
                sender.try_send(()),
                Err(TrySendError::Disconnected(()))
            ));
            result
        });
        assert_eq!(result, Err(ReporterError::Panicked("writer panic".into())));
    }

    #[test]
    fn dropping_reporter_closes_queue_and_scope_joins_writer() {
        let mut written = Vec::new();
        std::thread::scope(|scope| {
            let reporter = Reporter::spawn(scope, 2, |event| {
                written.push(event);
                Ok(())
            })
            .unwrap();
            reporter.send(1).unwrap();
            reporter.send(2).unwrap();
            drop(reporter);
        });
        assert_eq!(written, vec![1, 2]);
    }

    #[test]
    fn capacity_is_always_bounded_and_nonzero() {
        std::thread::scope(|scope| {
            for capacity in [0, MAX_CAPACITY + 1] {
                assert!(matches!(Reporter::<()>::spawn(scope, capacity, |_| Ok(())),
                    Err(ReporterError::InvalidCapacity(actual)) if actual == capacity));
            }
            Reporter::<()>::spawn(scope, DEFAULT_CAPACITY, |_| Ok(()))
                .unwrap()
                .finish()
                .unwrap();
        });
    }
}
