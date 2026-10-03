//! Closure upvalue machinery.
//!
//! `capture_upvalue` promotes a stack slot to an `Upvalue` (shared with
//! any closure that captures the same slot). `close_upvalues` copies the
//! stack value into the Upvalue's `closed` slot when the enclosing frame
//! returns, so closures can outlive their defining frame.

use crate::value::{Upvalue, UpvalueLocation, Value};
use crate::vm::VM;
use std::sync::{Arc, Mutex};

impl VM {
    pub(crate) fn capture_upvalue(&mut self, stack_idx: usize) -> Arc<Mutex<Upvalue>> {
        for uv in &self.open_upvalues {
            if let UpvalueLocation::Open(idx) = uv.lock().unwrap().location {
                if idx == stack_idx {
                    return uv.clone();
                }
            }
        }
        let uv = Arc::new(Mutex::new(Upvalue {
            location: UpvalueLocation::Open(stack_idx),
        }));
        self.open_upvalues.push(uv.clone());
        uv
    }

    pub(crate) fn close_upvalues(&mut self, from: usize) {
        if self.open_upvalues.is_empty() {
            return;
        }
        let stack = &self.stack;
        self.open_upvalues.retain(|uv| {
            let mut upvalue = uv.lock().unwrap();
            if let UpvalueLocation::Open(idx) = upvalue.location {
                if idx >= from {
                    // Lazy-locals convention: an unwritten captured slot may
                    // lie beyond the materialized stack.
                    upvalue.location = UpvalueLocation::Closed(
                        stack.get(idx).cloned().unwrap_or(Value::Null),
                    );
                    return false;
                }
            }
            true
        });
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn closes_many_upvalues_in_one_pass_and_keeps_outer_slots() {
        let mut vm = VM::new();
        vm.stack = (0..20_000).map(Value::I32).collect();
        let upvalues: Vec<_> = (0..20_000)
            .map(|idx| Arc::new(Mutex::new(Upvalue {
                location: UpvalueLocation::Open(idx),
            })))
            .collect();
        vm.open_upvalues = upvalues.clone();

        vm.close_upvalues(10_000);

        assert_eq!(vm.open_upvalues.len(), 10_000);
        assert!(matches!(upvalues[9_999].lock().unwrap().location, UpvalueLocation::Open(9_999)));
        assert!(matches!(upvalues[10_000].lock().unwrap().location, UpvalueLocation::Closed(Value::I32(10_000))));
        assert!(matches!(upvalues[19_999].lock().unwrap().location, UpvalueLocation::Closed(Value::I32(19_999))));
    }
}
