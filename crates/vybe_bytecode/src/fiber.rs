//! Fiber — a suspendable execution context.
//! When `await` is hit, the current fiber is suspended (stack + frames saved).
//! When the awaited Promise resolves, the fiber is resumed.

use crate::value::{Value, Upvalue};
use std::sync::{Arc, Mutex};

/// A suspended execution context — everything needed to resume.
#[derive(Debug)]
pub struct Fiber {
    /// Saved operand stack.
    pub stack: Vec<Value>,
    /// Saved call frames.
    pub frames: Vec<SavedFrame>,
    /// Saved open upvalues.
    pub open_upvalues: Vec<Arc<Mutex<Upvalue>>>,
    /// Saved structured-control-flow label entries. Generators that
    /// yield inside a `while` / `for` / `block` must preserve the
    /// surrounding block/loop labels so `br_label` / `br_if_label`
    /// target the right depth after resumption.
    pub label_stack: Vec<crate::vm::LabelEntry>,
    /// Saved active-continuation stack. A fiber captured inside a
    /// running coroutine (e.g. `await` inside a `@generator`) must
    /// bring its coroutine-context with it; restoring the fiber
    /// resurrects the nested generator state along with everything
    /// else.
    pub active_continuations: Vec<crate::vm::ActiveContinuation>,
    /// The value to push onto the stack when resuming (the await result).
    pub resume_value: Option<Value>,
}

#[derive(Debug, Clone)]
pub struct SavedFrame {
    pub chunk_index: usize,
    pub ip: usize,
    pub base: usize,
    pub upvalues: Vec<Arc<Mutex<Upvalue>>>,
}

impl Fiber {
    pub fn new(stack: Vec<Value>, frames: Vec<SavedFrame>, open_upvalues: Vec<Arc<Mutex<Upvalue>>>) -> Self {
        Fiber {
            stack,
            frames,
            open_upvalues,
            label_stack: Vec::new(),
            active_continuations: Vec::new(),
            resume_value: None,
        }
    }
    pub fn with_labels(mut self, labels: Vec<crate::vm::LabelEntry>) -> Self {
        self.label_stack = labels;
        self
    }
    pub fn with_continuations(mut self, conts: Vec<crate::vm::ActiveContinuation>) -> Self {
        self.active_continuations = conts;
        self
    }
}
