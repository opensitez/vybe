use crate::{VM, VMError};
use crate::component::ValType;
use crate::canon_layout::MemLayout;
use crate::shared_memory::SharedMemory;
use std::sync::Arc;

#[derive(Clone)]
pub(crate) struct GuestStreamWrite {
    pub memory: SharedMemory,
    pub ptr: usize,
    pub remaining: usize,
    pub element: Option<(Arc<ValType>, Arc<MemLayout>)>,
}

impl VM {
    pub(crate) fn complete_guest_host_stream_writes(&mut self) -> Result<bool, VMError> {
        let mut progressed = false;
        loop {
            let id = self.event_loop.borrow_mut().pending_guest_consumers.pop_front();
            let Some(id) = id else { break };
            if self.event_loop.borrow().has_stream_consumer(id) {
                if let Some(write) = self.guest_stream_writes.get(&id) {
                    progressed |= write.remaining != 0;
                    self.materialize_guest_stream_write(id, write.remaining)?;
                }
            }
        }
        Ok(progressed)
    }

    pub(crate) fn materialize_guest_stream_write(&mut self, id: u64, capacity: usize) -> Result<(), VMError> {
        if self.event_loop.borrow().stream_has_item(id) { return Ok(()) }
        let Some(write) = self.guest_stream_writes.get(&id) else { return Ok(()) };
        let count = capacity.min(write.remaining);
        if count == 0 { return Ok(()) }
        if let Some((element, layout)) = &write.element {
            let stride = layout.size as usize;
            let mut values = Vec::with_capacity(count);
            for index in 0..count {
                values.push(crate::canon_value::load_with_layout(
                    &write.memory, element, (write.ptr + index * layout.size as usize) as u32, layout,
                ).map_err(|error| VMError::new(format!("canon stream.write: {error}")))?);
            }
            let write = self.guest_stream_writes.get_mut(&id).unwrap();
            write.ptr += count * stride;
            write.remaining -= count;
            let mut events = self.event_loop.borrow_mut();
            for value in values {
                if let Some(fiber) = events.stream_push(id, value).map_err(VMError::new)? {
                    events.immediate.push_back(crate::event_loop::Task::ResumeFiber(fiber));
                }
            }
        } else {
            let mut bytes = vec![0; count];
            write.memory.read_bytes(write.ptr, &mut bytes);
            let write = self.guest_stream_writes.get_mut(&id).unwrap();
            write.ptr += count;
            write.remaining -= count;
            self.event_loop.borrow_mut().stream_push_bytes(id, bytes);
        }
        Ok(())
    }

    pub(crate) fn retain_guest_stream_write(
        &mut self, handle: u32, id: u64, write: GuestStreamWrite,
        memory: u32, realloc: Option<u32>,
    ) -> Result<(), VMError> {
        let ptr = write.ptr;
        let n = write.remaining;
        self.guest_stream_writes.insert(id, write);
        let capacity = self.async_stream_reads.get(&id).map(|read| read.n)
            .or_else(|| self.event_loop.borrow().pending_stream_read_capacity(id));
        self.event_loop.borrow_mut().begin_guest_stream_write(id, n);
        if let Some(capacity) = capacity {
            self.materialize_guest_stream_write(id, capacity)?;
            self.complete_async_stream_read(id)?;
        }
        let result = self.event_loop.borrow_mut().take_stream_zero_completion(id, false);
        if let Some(result) = result {
            self.guest_stream_writes.remove(&id);
            self.push(crate::Value::I32(result as i32))?;
            return Ok(());
        }
        if let Some(crate::handle_table::HandleEntry::WritableStreamEnd(end)) = self.handle_table.get_mut(handle) {
            end.state = crate::handle_table::CopyState::Copying;
        }
        if self.canon_async_opt() {
            self.push(crate::Value::I32(crate::canon_copy::BLOCKED as i32))?;
            Ok(())
        } else {
            Err(self.park_sync_copy(crate::fiber::PendingCopy {
                handle, end_id: id, ptr, n, memory, realloc,
                kind: crate::fiber::PendingCopyKind::StreamWritePending,
            }))
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::canon_def::{CanonDef, CanonOpts};
    use crate::vm::CanonBuiltin;
    use crate::Value;

    #[test]
    #[ignore = "pending guest-buffer cost diagnostic, not an application speedup"]
    fn bench_pending_write_and_cancel_do_not_copy_the_payload() {
        let mut vm = VM::new();
        vm.memory.resize(16 * 1024 * 1024, 0);
        vm.canon_types = vec![Some(ValType::Stream(Box::new(ValType::U8)))];
        vm.canon_defs = vec![CanonDef {
            ty: Some(0),
            opts: CanonOpts { memory: Some(0), is_async: true, ..CanonOpts::default() },
            ..CanonDef::default()
        }];
        vm.canon_type_immediate = Some(0);
        vm.exec_canon_builtin(CanonBuiltin::StreamNew).unwrap();
        let pair = vm.pop().as_i64() as u64;
        let writer = (pair >> 32) as u32;
        let crate::handle_table::HandleEntry::WritableStreamEnd(end) =
            vm.handle_table.get(writer).unwrap() else { unreachable!() };
        let id = end.id;
        for size in [1, 256 * 1024, 16 * 1024 * 1024] {
            let start = std::time::Instant::now();
            for _ in 0..2000 {
                for value in [writer as i32, 0, size] { vm.push(Value::I32(value)).unwrap(); }
                vm.exec_canon_builtin(CanonBuiltin::StreamWrite).unwrap();
                assert_eq!(vm.pop().as_i32(), crate::canon_copy::BLOCKED as i32);
                assert!(!vm.event_loop.borrow().stream_has_item(id));
                vm.push(Value::I32(writer as i32)).unwrap();
                vm.exec_canon_builtin(CanonBuiltin::StreamCancelWrite).unwrap();
                assert_eq!(vm.pop().as_i32(), crate::canon_copy::CopyResult::Cancelled as i32);
            }
            assert!(vm.guest_stream_writes.is_empty());
            eprintln!("{size} bytes pending write+cancel: {} ns/pair", start.elapsed().as_nanos() / 2000);
        }
    }
}
