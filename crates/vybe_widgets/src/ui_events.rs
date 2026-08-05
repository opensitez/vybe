//! W3C UI Events — the queue and the pointer state derived from it.
//!
//! THE QUEUE LIVES HERE because this crate is the web platform's
//! implementation: the same reason the document does. A native window backend
//! (winit) pushes events in; in a browser the real DOM would fill the same
//! queue from `addEventListener`, and nothing above changes.
//!
//! This is deliberately separate from [`layout::MouseEvent`](crate::layout)
//! and `KeyEvent`, which are the toolkit's own internal dispatch types. These
//! carry the spec's field names and value conventions — `type`, `key`, `code`,
//! `clientX`, `clientY`, `button`, `buttons`, `deltaY`, `ctrlKey`… — so a
//! guest that knows the DOM needs no translation. Adapters speaking another
//! vocabulary (SDL maps `mousedown` → `SDL_MOUSEBUTTONDOWN`, DOM's 0-based
//! `button` → SDL's 1-based) translate on THEIR side, which is what keeps
//! this one standard.
//!
//! Nothing here touches a VM value: building the guest-visible event object
//! is the host's job, not the toolkit's.

use std::collections::VecDeque;
use std::sync::{Arc, Mutex, OnceLock};

/// One UI event in W3C shape. Field names match the IDL attributes.
#[derive(Clone, Debug, Default)]
pub struct UiEvent {
    /// `keydown` | `keyup` | `mousedown` | `mouseup` | `mousemove` | `wheel`
    pub kind: String,
    /// `KeyboardEvent.key` — the character or named key ("a", "Enter",
    /// "ArrowLeft", "+"). Empty for mouse events.
    pub key: String,
    /// `KeyboardEvent.code` — the physical key ("KeyA", "Digit1",
    /// "ArrowLeft"), layout-independent.
    pub code: String,
    /// Legacy `keyCode`. Deprecated in the DOM, kept because adapters
    /// (SDL keysyms) want a numeric identity and the browser still ships it.
    pub key_code: i32,
    pub client_x: i32,
    pub client_y: i32,
    /// `MouseEvent.button` — 0 left, 1 middle, 2 right (DOM numbering).
    pub button: i32,
    /// `MouseEvent.buttons` — bitmask: 1 left, 2 right, 4 middle.
    pub buttons: i32,
    pub delta_y: f64,
    pub ctrl_key: bool,
    pub shift_key: bool,
    pub alt_key: bool,
    pub meta_key: bool,
    /// `KeyboardEvent.repeat` — held-key auto-repeat.
    pub repeat: bool,
}

/// Pointer/modifier state, as a browser tracks it between events.
#[derive(Clone, Copy, Debug, Default)]
pub struct PointerState {
    pub client_x: i32,
    pub client_y: i32,
    pub buttons: i32,
    pub ctrl_key: bool,
    pub shift_key: bool,
    pub alt_key: bool,
    pub meta_key: bool,
}

/// Bound so a stuck producer can never exhaust memory; the browser drops
/// events under pressure too.
const MAX_QUEUED: usize = 4096;

/// The UI-event queue plus the pointer/modifier state derived from it.
pub struct UiEventQueue {
    events: Mutex<VecDeque<UiEvent>>,
    pointer: Mutex<PointerState>,
}

impl UiEventQueue {
    fn new() -> Self {
        UiEventQueue {
            events: Mutex::new(VecDeque::new()),
            pointer: Mutex::new(PointerState::default()) }
    }

    /// Enqueue an event, updating the tracked pointer/modifier state the way
    /// a browser does as events flow.
    pub fn push(&self, evt: UiEvent) {
        {
            let mut st = self.pointer.lock().unwrap();
            st.ctrl_key = evt.ctrl_key;
            st.shift_key = evt.shift_key;
            st.alt_key = evt.alt_key;
            st.meta_key = evt.meta_key;
            match evt.kind.as_str() {
                "mousemove" | "mousedown" | "mouseup" => {
                    st.client_x = evt.client_x;
                    st.client_y = evt.client_y;
                    st.buttons = evt.buttons;
                }
                _ => {}
            }
        }
        let mut q = self.events.lock().unwrap();
        if q.len() >= MAX_QUEUED {
            q.pop_front();
        }
        q.push_back(evt);
    }

    /// Dequeue the oldest event.
    pub fn poll(&self) -> Option<UiEvent> {
        self.events.lock().unwrap().pop_front()
    }

    pub fn pointer_state(&self) -> PointerState {
        *self.pointer.lock().unwrap()
    }

    pub fn pending(&self) -> usize {
        self.events.lock().unwrap().len()
    }
}

/// The process-wide queue. A native window backend and the guest-facing host
/// functions must see the SAME queue, so it is reachable without a handle.
pub fn queue() -> Arc<UiEventQueue> {
    static QUEUE: OnceLock<Arc<UiEventQueue>> = OnceLock::new();
    QUEUE.get_or_init(|| Arc::new(UiEventQueue::new())).clone()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn pointer_state_tracks_the_event_stream() {
        let q = UiEventQueue::new();
        q.push(UiEvent {
            kind: "mousemove".into(),
            client_x: 12,
            client_y: 34,
            buttons: 1,
            shift_key: true,
            ..UiEvent::default()
        });
        let st = q.pointer_state();
        assert_eq!((st.client_x, st.client_y, st.buttons), (12, 34, 1));
        assert!(st.shift_key);
        // A key event updates modifiers but must NOT move the pointer.
        q.push(UiEvent {
            kind: "keydown".into(),
            shift_key: false,
            ..UiEvent::default()
        });
        let st = q.pointer_state();
        assert_eq!((st.client_x, st.client_y), (12, 34));
        assert!(!st.shift_key);
    }

    #[test]
    fn the_queue_is_fifo_and_bounded() {
        let q = UiEventQueue::new();
        for i in 0..(MAX_QUEUED + 10) {
            q.push(UiEvent {
                kind: "keydown".into(),
                key_code: i as i32,
                ..UiEvent::default()
            });
        }
        assert_eq!(q.pending(), MAX_QUEUED, "the queue must stay bounded");
        // The OLDEST were dropped, so the front is now event #10.
        assert_eq!(q.poll().unwrap().key_code, 10);
    }
}
