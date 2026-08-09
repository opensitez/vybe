//! Deadline scheduling — HTML timers and the animation frame clock.
//!
//! A toolkit needs both on its own: a frame clock to animate, and timers for
//! a blinking cursor, a tooltip delay, a debounce. So the MECHANISM lives
//! here — when things are due, in what order, and what cancelling means.
//!
//! **Nothing here stores a callback.** These deal in ids, exactly as the DOM
//! side deals in `WidgetEvent`s: the toolkit says *what became due*, and
//! whoever registered the work decides what running it means. That is what
//! keeps this crate free of any runtime's value type, and it is the same
//! shape the widget tree already had — record, don't call back.
//!
//! Deadlines are reported as a RELATIVE delay, never an absolute timestamp.
//! A host runs its own monotonic clock with its own origin; handing it an
//! absolute time from this clock would silently mis-schedule every sleep.

use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Mutex, OnceLock};
use std::time::Instant;

/// Milliseconds since this module was first touched. Monotonic, with an
/// origin private to this crate — which is exactly why deadlines cross the
/// boundary as relative delays.
pub fn now_ms() -> f64 {
    static ORIGIN: OnceLock<Instant> = OnceLock::new();
    ORIGIN.get_or_init(Instant::now).elapsed().as_secs_f64() * 1000.0
}

struct Entry {
    id: u64,
    fire_at_ms: f64,
}

/// The HTML timer wheel, minus the callbacks.
///
/// Entries stay in REGISTRATION order and fire first-registered-due-first, so
/// `setTimeout(f, 0); setTimeout(g, 0)` runs `f` before `g`.
pub struct Timers {
    entries: Mutex<Vec<Entry>>,
    next_id: AtomicU64,
}

impl Timers {
    fn new() -> Self {
        Timers {
            entries: Mutex::new(Vec::new()),
            next_id: AtomicU64::new(1),
        }
    }

    /// Schedule `delay_ms` from now; returns the cancellable id.
    pub fn schedule(&self, delay_ms: f64) -> u64 {
        let id = self.next_id.fetch_add(1, Ordering::Relaxed);
        self.entries.lock().unwrap().push(Entry {
            id,
            fire_at_ms: now_ms() + delay_ms.max(0.0),
        });
        id
    }

    /// Cancel by id. True if it was still queued.
    pub fn cancel(&self, id: u64) -> bool {
        let mut entries = self.entries.lock().unwrap();
        match entries.iter().position(|e| e.id == id) {
            Some(pos) => {
                entries.remove(pos);
                true
            }
            None => false,
        }
    }

    /// Pop ONE due id — first-registered-due-first, matching a drain loop's
    /// one-task-per-turn contract. `None` if nothing is due yet.
    pub fn take_due(&self) -> Option<u64> {
        let now = now_ms();
        let mut entries = self.entries.lock().unwrap();
        let pos = entries.iter().position(|e| e.fire_at_ms <= now)?;
        Some(entries.remove(pos).id)
    }

    /// How long until the earliest deadline, in ms; `0.0` when one is already
    /// due. `None` when nothing is scheduled.
    pub fn delay_until_next_ms(&self) -> Option<f64> {
        let now = now_ms();
        self.entries
            .lock()
            .unwrap()
            .iter()
            .map(|e| (e.fire_at_ms - now).max(0.0))
            .reduce(f64::min)
    }

    pub fn pending(&self) -> usize {
        self.entries.lock().unwrap().len()
    }

    /// Every queued id, in registration order.
    pub fn ids(&self) -> Vec<u64> {
        self.entries.lock().unwrap().iter().map(|e| e.id).collect()
    }
}

/// 60 Hz. The spec says "before the next repaint" and leaves the rate to the
/// display; a fixed cadence is the honest approximation for a backend with no
/// vsync signal to hand us.
const FRAME_MS: f64 = 1000.0 / 60.0;

/// `requestAnimationFrame`, minus the callbacks.
///
/// Per spec a registration fires AT MOST once; a caller wanting the next
/// frame registers again. That is what makes "stop drawing when nobody asks"
/// fall out for free.
pub struct Frames {
    pending: Mutex<Vec<u64>>,
    next_frame_ms: Mutex<f64>,
    next_id: AtomicU64,
}

impl Frames {
    fn new() -> Self {
        Frames {
            pending: Mutex::new(Vec::new()),
            next_frame_ms: Mutex::new(0.0),
            next_id: AtomicU64::new(1),
        }
    }

    /// Register for the next frame; returns the cancellable id.
    pub fn request(&self) -> u64 {
        let id = self.next_id.fetch_add(1, Ordering::Relaxed);
        {
            let mut next = self.next_frame_ms.lock().unwrap();
            // The first request after an idle period draws promptly rather
            // than waiting out a frame nobody was rendering.
            if *next <= 0.0 {
                *next = now_ms();
            }
        }
        self.pending.lock().unwrap().push(id);
        id
    }

    pub fn cancel(&self, id: u64) -> bool {
        let mut pending = self.pending.lock().unwrap();
        match pending.iter().position(|p| *p == id) {
            Some(pos) => {
                pending.remove(pos);
                true
            }
            None => false,
        }
    }

    pub fn take_due(&self) -> Option<u64> {
        let now = now_ms();
        if now < *self.next_frame_ms.lock().unwrap() {
            return None;
        }
        let mut pending = self.pending.lock().unwrap();
        if pending.is_empty() {
            return None;
        }
        let id = pending.remove(0);
        // The frame advances once its LAST registration has been handed out,
        // so everything registered for a given frame runs in THAT frame
        // rather than being spread across several.
        if pending.is_empty() {
            *self.next_frame_ms.lock().unwrap() = now + FRAME_MS;
        }
        Some(id)
    }

    pub fn delay_until_next_ms(&self) -> Option<f64> {
        if self.pending.lock().unwrap().is_empty() {
            return None;
        }
        Some((*self.next_frame_ms.lock().unwrap() - now_ms()).max(0.0))
    }

    pub fn pending(&self) -> usize {
        self.pending.lock().unwrap().len()
    }

    pub fn ids(&self) -> Vec<u64> {
        self.pending.lock().unwrap().clone()
    }
}

/// The process-wide wheel. A host and any in-process caller must see the SAME
/// one, so it is reachable without a handle.
pub fn timers() -> Arc<Timers> {
    static TIMERS: OnceLock<Arc<Timers>> = OnceLock::new();
    TIMERS.get_or_init(|| Arc::new(Timers::new())).clone()
}

/// Drop every scheduled timer and pending frame.
///
/// Both wheels are process-wide but their entries belong to whichever program
/// scheduled them. A VM that serves more than one program (the warm worker,
/// `--serve`) must not let one program's pending work fire during another's,
/// so this is called from the owning plugin's `reset`.
pub fn reset() {
    timers().entries.lock().unwrap().clear();
    frames().pending.lock().unwrap().clear();
}

/// The process-wide frame clock.
pub fn frames() -> Arc<Frames> {
    static FRAMES: OnceLock<Arc<Frames>> = OnceLock::new();
    FRAMES.get_or_init(|| Arc::new(Frames::new())).clone()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_zero_delay_timer_is_due_immediately_and_in_order() {
        let t = Timers::new();
        let f = t.schedule(0.0);
        let g = t.schedule(0.0);
        assert_eq!(t.delay_until_next_ms(), Some(0.0));
        // First registered, first out — `setTimeout(f,0); setTimeout(g,0)`.
        assert_eq!(t.take_due(), Some(f));
        assert_eq!(t.take_due(), Some(g));
        assert_eq!(t.take_due(), None);
        assert_eq!(t.pending(), 0);
    }

    #[test]
    fn a_future_timer_is_not_due_and_reports_a_relative_delay() {
        let t = Timers::new();
        t.schedule(50_000.0);
        assert_eq!(t.take_due(), None, "not due yet");
        let delay = t.delay_until_next_ms().expect("a deadline");
        assert!(
            delay > 49_000.0 && delay <= 50_000.0,
            "delay must be relative, got {delay}"
        );
    }

    #[test]
    fn cancelling_removes_the_deadline() {
        let t = Timers::new();
        let id = t.schedule(0.0);
        assert!(t.cancel(id));
        assert!(!t.cancel(id), "cancelling twice reports nothing was queued");
        assert_eq!(t.take_due(), None);
        assert_eq!(t.delay_until_next_ms(), None);
    }

    #[test]
    fn every_registration_for_a_frame_runs_in_that_frame() {
        let f = Frames::new();
        let a = f.request();
        let b = f.request();
        assert_eq!(f.take_due(), Some(a));
        assert_eq!(f.take_due(), Some(b), "same frame, not the next one");
        // The frame has advanced, so a fresh request waits.
        f.request();
        assert_eq!(f.take_due(), None, "the next frame is not here yet");
        let delay = f.delay_until_next_ms().expect("a frame deadline");
        assert!(delay > 0.0 && delay <= FRAME_MS, "got {delay}");
    }

    #[test]
    fn a_frame_registration_fires_at_most_once() {
        let f = Frames::new();
        let id = f.request();
        assert_eq!(f.take_due(), Some(id));
        assert_eq!(f.pending(), 0, "no re-arming; the caller re-registers");
    }
}
