use vybe_runtime::{Chunk, Op};

/// Copy a marshalled byte buffer, retrying short synchronous canonical copies.
/// Empty payloads need only the caller's EOF; non-COMPLETED ends cannot retry.
pub(super) fn emit_write_buffer(
    chunk: &mut Chunk, writer: u16, ptr: u16, remaining: u16,
    write_import: u16, line: u32,
) {
    let result = chunk.alloc_scratch(1);
    let copied = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, remaining, line);
    chunk.emit_if(line);
    let done = chunk.emit_block(line);
    let (write, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, writer, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ptr, line);
    chunk.emit_op_u16(Op::LOCAL_GET, remaining, line);
    chunk.emit_call(write_import, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    chunk.emit_i32_const(0xf, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_br_if(1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    chunk.emit_i32_const(4, line);
    chunk.emit_op(Op::I32_SHR_U, line);
    chunk.emit_op_u16(Op::LOCAL_SET, copied, line);
    // Complete host-buffer copies take the exit without updating the cursor.
    chunk.emit_op_u16(Op::LOCAL_GET, copied, line);
    chunk.emit_op_u16(Op::LOCAL_GET, remaining, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_br_if(1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ptr, line);
    chunk.emit_op_u16(Op::LOCAL_GET, copied, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ptr, line);
    chunk.emit_op_u16(Op::LOCAL_GET, remaining, line);
    chunk.emit_op_u16(Op::LOCAL_GET, copied, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, remaining, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(write);
    chunk.emit_end(line);
    chunk.patch_block(done);
    chunk.emit_end(line);
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::{Arc, Mutex};
    use vybe_runtime::Value;

    fn fixture(size: i32) -> Chunk {
        let mut chunk = Chunk::new("canonical-buffer-progress");
        let write = chunk.add_import("copy-contract-test", "write");
        let writer = chunk.alloc_scratch(1);
        let ptr = chunk.alloc_scratch(1);
        let remaining = chunk.alloc_scratch(1);
        for (slot, value) in [(writer, 7), (ptr, 4096), (remaining, size)] {
            chunk.emit_i32_const(value, 0);
            chunk.emit_op_u16(Op::LOCAL_SET, slot, 0);
        }
        emit_write_buffer(&mut chunk, writer, ptr, remaining, write, 0);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::RETURN, 0);
        chunk
    }

    #[test]
    fn retries_only_the_uncopied_suffix_and_skips_empty_payloads() {
        for size in [0, 5] {
            let calls = Arc::new(Mutex::new(Vec::new()));
            let captured = calls.clone();
            let mut vm = vybe_runtime::VM::new();
            vm.register_host_fn("copy-contract-test", "write", Box::new(move |_, args| {
                assert_eq!(args[0].as_i32(), 7);
                let (ptr, count) = (args[1].as_i32(), args[2].as_i32());
                captured.lock().unwrap().push((ptr, count));
                Value::I32(count.min(2) << 4)
            }));
            vm.run(vec![fixture(size)]).unwrap();
            assert_eq!(&*calls.lock().unwrap(), if size == 0 { &[][..] }
                else { &[(4096, 5), (4098, 3), (4100, 1)][..] });
        }
    }

    #[test]
    fn dropped_end_is_not_retried() {
        let calls = Arc::new(Mutex::new(0));
        let captured = calls.clone();
        let mut vm = vybe_runtime::VM::new();
        vm.register_host_fn("copy-contract-test", "write", Box::new(move |_, _| {
            *captured.lock().unwrap() += 1;
            Value::I32((2 << 4) | 1)
        }));
        vm.run(vec![fixture(5)]).unwrap();
        assert_eq!(*calls.lock().unwrap(), 1);
    }

    #[test]
    fn progress_loop_is_standard_wasm() {
        let binary = vybe_platform_wasm::write_wasm(&[fixture(5)]);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary).expect("canonical progress loop must validate");
    }
}
