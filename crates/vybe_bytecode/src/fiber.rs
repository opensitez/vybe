//! Fiber — a suspendable execution context.
//! When `await` is hit, the current fiber is suspended (stack + frames saved).
//! When the awaited Promise resolves, the fiber is resumed.

use crate::value::Value;
use crate::value::{Upvalue, UpvalueLocation};
use std::cell::RefCell;
use std::rc::Rc;

/// A suspended execution context — everything needed to resume.
#[derive(Debug)]
pub struct Fiber {
    /// Saved operand stack.
    pub stack: Vec<Value>,
    /// Saved call frames.
    pub frames: Vec<SavedFrame>,
    /// Saved open upvalues.
    pub open_upvalues: Vec<Rc<RefCell<Upvalue>>>,
    /// The value to push onto the stack when resuming (the await result).
    pub resume_value: Option<Value>,
}

#[derive(Debug, Clone)]
pub struct SavedFrame {
    pub chunk_index: usize,
    pub ip: usize,
    pub base: usize,
    pub upvalues: Vec<Rc<RefCell<Upvalue>>>,
}

impl Fiber {
    pub fn new(stack: Vec<Value>, frames: Vec<SavedFrame>, open_upvalues: Vec<Rc<RefCell<Upvalue>>>) -> Self {
        Fiber {
            stack,
            frames,
            open_upvalues,
            resume_value: None,
        }
    }
}
