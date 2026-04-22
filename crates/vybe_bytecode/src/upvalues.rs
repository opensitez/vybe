//! Closure upvalue machinery.
//!
//! `capture_upvalue` promotes a stack slot to an `Upvalue` (shared with
//! any closure that captures the same slot). `close_upvalues` copies the
//! stack value into the Upvalue's `closed` slot when the enclosing frame
//! returns, so closures can outlive their defining frame.

use std::sync::{Arc, Mutex};
use crate::value::{Upvalue, UpvalueLocation};
use crate::vm::VM;

impl VM {
    pub(crate) fn capture_upvalue(&mut self, stack_idx: usize) -> Arc<Mutex<Upvalue>> {
        for uv in &self.open_upvalues {
            if let UpvalueLocation::Open(idx) = uv.lock().unwrap().location {
                if idx == stack_idx { return uv.clone(); }
            }
        }
        let uv = Arc::new(Mutex::new(Upvalue { location: UpvalueLocation::Open(stack_idx) }));
        self.open_upvalues.push(uv.clone());
        uv
    }

    pub(crate) fn close_upvalues(&mut self, from: usize) {
        let mut i = 0;
        while i < self.open_upvalues.len() {
            let should_close = matches!(
                self.open_upvalues[i].lock().unwrap().location,
                UpvalueLocation::Open(idx) if idx >= from
            );
            if should_close {
                let uv = self.open_upvalues.remove(i);
                let mut u = uv.lock().unwrap();
                if let UpvalueLocation::Open(idx) = u.location {
                    u.location = UpvalueLocation::Closed(self.stack[idx].clone());
                }
            } else {
                i += 1;
            }
        }
    }
}
