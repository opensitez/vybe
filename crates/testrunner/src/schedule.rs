//! One shared worker queue with rolling suite admission and round-robin lanes.

use std::collections::{HashMap, VecDeque};
use std::path::{Path, PathBuf};
use std::sync::{Arc, Condvar, Mutex};

struct Lane {
    pending: VecDeque<PathBuf>,
    remaining: usize,
}

struct State {
    lanes: Vec<Lane>,
    active: VecDeque<usize>,
    waiting: VecDeque<usize>,
    in_flight: HashMap<PathBuf, usize>,
    remaining: usize,
    cancelled: bool,
}

pub struct Schedule {
    state: Mutex<State>,
    changed: Condvar,
    limit: usize,
    total: usize,
}

impl Schedule {
    pub fn new(groups: Vec<Vec<PathBuf>>, max_active: usize) -> Self {
        assert!(max_active > 0, "suite admission limit must be positive");
        let mut seen = std::collections::HashSet::new();
        let mut waiting = VecDeque::new();
        let mut total = 0;
        let lanes = groups
            .into_iter()
            .enumerate()
            .map(|(index, paths)| {
                for path in &paths {
                    assert!(seen.insert(path.clone()), "scheduled paths must be unique");
                }
                let remaining = paths.len();
                total += remaining;
                if remaining != 0 {
                    waiting.push_back(index);
                }
                Lane {
                    pending: paths.into(),
                    remaining,
                }
            })
            .collect();
        let schedule = Self {
            state: Mutex::new(State {
                lanes,
                active: VecDeque::new(),
                waiting,
                in_flight: HashMap::new(),
                remaining: total,
                cancelled: false,
            }),
            changed: Condvar::new(),
            limit: max_active,
            total,
        };
        schedule.admit(&mut schedule.state.lock().unwrap());
        schedule
    }

    fn admit(&self, state: &mut State) {
        while state.active.len() < self.limit {
            let Some(index) = state.waiting.pop_front() else {
                break;
            };
            state.active.push_back(index);
        }
    }

    /// Wait only when admitted lanes have in-flight work but no queued case.
    pub fn pop(&self) -> Option<PathBuf> {
        let mut state = self.state.lock().unwrap_or_else(|error| error.into_inner());
        loop {
            if state.cancelled || state.remaining == 0 {
                return None;
            }
            for _ in 0..state.active.len() {
                let index = state.active.pop_front().unwrap();
                state.active.push_back(index);
                if let Some(path) = state.lanes[index].pending.pop_front() {
                    state.in_flight.insert(path.clone(), index);
                    return Some(path);
                }
            }
            state = self
                .changed
                .wait(state)
                .unwrap_or_else(|error| error.into_inner());
        }
    }

    /// Call after the result callback, never merely when a case is dequeued.
    pub fn complete(&self, path: &Path) -> Option<usize> {
        let mut state = self.state.lock().unwrap_or_else(|error| error.into_inner());
        let index = state.in_flight.remove(path)?;
        state.remaining -= 1;
        state.lanes[index].remaining -= 1;
        let finished = state.lanes[index].remaining == 0;
        if finished {
            state.active.retain(|active| *active != index);
            self.admit(&mut state);
        }
        self.changed.notify_all();
        finished.then_some(index)
    }

    pub fn retry(&self, path: PathBuf) {
        let mut state = self.state.lock().unwrap_or_else(|error| error.into_inner());
        if let Some(index) = state.in_flight.remove(&path) {
            state.lanes[index].pending.push_front(path);
            self.changed.notify_all();
        }
    }

    pub fn cancel(&self) {
        self.state
            .lock()
            .unwrap_or_else(|error| error.into_inner())
            .cancelled = true;
        self.changed.notify_all();
    }

    pub fn total(&self) -> usize {
        self.total
    }

    pub fn is_cancelled(&self) -> bool {
        self.state
            .lock()
            .unwrap_or_else(|error| error.into_inner())
            .cancelled
    }

    pub fn is_complete(&self) -> bool {
        self.state
            .lock()
            .unwrap_or_else(|error| error.into_inner())
            .remaining
            == 0
    }

    pub(crate) fn lease(self: &Arc<Self>) -> Option<Lease> {
        self.pop().map(|path| Lease {
            schedule: self.clone(),
            path: Some(path),
        })
    }
}

pub(crate) struct Lease {
    schedule: Arc<Schedule>,
    path: Option<PathBuf>,
}

impl Lease {
    pub(crate) fn path(&self) -> &Path {
        self.path.as_deref().unwrap()
    }

    pub(crate) fn complete(mut self) -> Option<usize> {
        self.schedule.complete(&self.path.take().unwrap())
    }

    pub(crate) fn retry(mut self) {
        self.schedule.retry(self.path.take().unwrap());
    }
}

impl Drop for Lease {
    fn drop(&mut self) {
        if let Some(path) = self.path.take() {
            if std::thread::panicking() {
                self.schedule.cancel();
            }
            self.schedule.retry(path);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::mpsc;
    use std::time::Duration;

    fn groups(groups: &[&[&str]]) -> Vec<Vec<PathBuf>> {
        groups
            .iter()
            .map(|group| group.iter().map(|path| PathBuf::from(*path)).collect())
            .collect()
    }

    #[test]
    fn fast_suite_releases_next_lane_while_first_suite_stays_in_flight() {
        let schedule = Schedule::new(groups(&[&["A1", "A2"], &["B1"], &["C1"]]), 2);
        let a1 = schedule.pop().unwrap();
        let b1 = schedule.pop().unwrap();
        let a2 = schedule.pop().unwrap();
        assert_eq!(
            (&*a1, &*b1, &*a2),
            (Path::new("A1"), Path::new("B1"), Path::new("A2"))
        );
        assert_eq!(schedule.complete(&b1), Some(1));
        let c1 = schedule.pop().unwrap();
        assert_eq!(c1, Path::new("C1"));
        assert_eq!(schedule.complete(&c1), Some(2));
        assert_eq!(schedule.complete(&a2), None);
        assert_eq!(schedule.complete(&a1), Some(0));
        assert!(schedule.pop().is_none());
    }

    #[test]
    fn waiting_pop_wakes_on_completion_but_not_on_last_dispatch() {
        let schedule = Arc::new(Schedule::new(groups(&[&["A"], &["B"]]), 1));
        let a = schedule.pop().unwrap();
        let (tx, rx) = mpsc::channel();
        let copy = schedule.clone();
        let thread = std::thread::spawn(move || {
            let _ = tx.send(copy.pop());
        });
        let premature = rx.recv_timeout(Duration::from_millis(20));
        schedule.complete(&a);
        let b = rx.recv_timeout(Duration::from_secs(2));
        schedule.cancel();
        thread.join().unwrap();
        assert!(premature.is_err());
        assert_eq!(b.unwrap(), Some(PathBuf::from("B")));
    }

    #[test]
    fn retry_keeps_suite_slot_and_empty_lanes_keep_original_indices() {
        let schedule = Schedule::new(groups(&[&[], &["A"], &[], &["B"]]), 1);
        let a = schedule.pop().unwrap();
        schedule.retry(a.clone());
        assert_eq!(schedule.pop(), Some(a.clone()));
        assert_eq!(schedule.complete(&a), Some(1));
        assert_eq!(schedule.complete(&a), None);
        let b = schedule.pop().unwrap();
        assert_eq!(schedule.complete(&b), Some(3));
        assert_eq!(schedule.total(), 2);
        assert!(schedule.is_complete());
    }

    #[test]
    fn panicked_lease_cancels_waiters_without_marking_case_complete() {
        let schedule = Arc::new(Schedule::new(groups(&[&["A"], &["B"]]), 1));
        let lease = schedule.lease().unwrap();
        let (tx, rx) = mpsc::channel();
        let copy = schedule.clone();
        let waiter = std::thread::spawn(move || {
            let _ = tx.send(copy.pop());
        });
        let panic = std::panic::catch_unwind(std::panic::AssertUnwindSafe(move || {
            let _lease = lease;
            panic!("synthetic worker panic");
        }));
        let result = rx.recv_timeout(Duration::from_secs(2));
        schedule.cancel();
        waiter.join().unwrap();
        assert!(panic.is_err());
        assert_eq!(result.unwrap(), None);
        assert!(schedule.is_cancelled());
        assert!(!schedule.is_complete());
    }

    #[test]
    fn abandoned_nonpanicking_lease_requeues_instead_of_losing_case() {
        let schedule = Arc::new(Schedule::new(groups(&[&["A"]]), 1));
        drop(schedule.lease().unwrap());
        let lease = schedule.lease().unwrap();
        assert_eq!(lease.path(), Path::new("A"));
        lease.complete();
        assert!(schedule.is_complete());
    }
}
