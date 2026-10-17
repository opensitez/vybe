//! Branch payloads follow their target label, including across void guards.
use vybe_runtime::{Chunk, Op};

fn equal(chunk: &mut Chunk) {
    chunk.emit_f64_const(3.0, 1);
    chunk.emit_f64_const(3.0, 1);
    chunk.emit_op(Op::F64_EQ, 1);
}

fn validate(chunk: Chunk) {
    let name = chunk.name.clone();
    let bytes = vybe_platform_wasm::write_wasm(&[chunk]);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    export(&name, &bytes);
}

fn export(name: &str, bytes: &[u8]) {
    if let Some(directory) = std::env::var_os("VYBE_I32_BRANCH_EXPORT_DIR") {
        let directory = std::path::PathBuf::from(directory);
        std::fs::create_dir_all(&directory).unwrap();
        std::fs::write(directory.join(format!("{name}.wasm")), bytes).unwrap();
    }
}

#[test]
fn raw_i32_branch_crosses_void_guard() {
    let mut c = Chunk::new("raw_i32_branch");
    c.emit_block_i32(1);
    c.emit_i32_const(1, 1);
    c.emit_if(1);
    equal(&mut c);
    c.emit_br(1, 1);
    c.emit_end(1);
    c.emit_i32_const(0, 1);
    c.emit_end(1);
    c.emit_op(Op::RETURN, 1);
    validate(c);
}

#[test]
fn raw_i32_join_result_branches_across_two_void_guards() {
    let mut c = Chunk::new("raw_i32_join_branch");
    c.emit_block_i32(1);
    c.emit_i32_const(1, 1);
    c.emit_if(1);
    c.emit_i32_const(1, 1);
    c.emit_if(1);
    c.emit_i32_const(1, 1);
    c.emit_if_i32(1);
    equal(&mut c);
    c.emit_else(1);
    c.emit_i32_const(0, 1);
    c.emit_end(1);
    c.emit_br(2, 1);
    c.emit_end(1);
    c.emit_end(1);
    c.emit_i32_const(0, 1);
    c.emit_end(1);
    c.emit_op(Op::RETURN, 1);
    validate(c);
}

#[test]
fn branch_to_value_label_still_boxes_i32_join() {
    let mut c = Chunk::new("boxed_i32_join_branch");
    c.emit_block_typed(1, 1);
    c.emit_i32_const(1, 1);
    c.emit_if_i32(1);
    c.emit_i32_const(7, 1);
    c.emit_else(1);
    c.emit_i32_const(8, 1);
    c.emit_end(1);
    c.emit_br(0, 1);
    c.emit_end(1);
    c.emit_op(Op::RETURN, 1);
    validate(c);
}

#[test]
fn raw_i32_branch_counts_try_table_label() {
    let mut c = Chunk::new("raw_i32_try_branch");
    c.emit_block_i32(1);
    c.emit_try_table_clauses(0, 0, &[], 1);
    equal(&mut c);
    c.emit_br(1, 1);
    c.emit_end(1);
    c.emit_i32_const(0, 1);
    c.emit_end(1);
    c.emit_op(Op::RETURN, 1);
    validate(c);
}

#[test]
fn raw_i32_branch_to_implicit_function_label() {
    let mut c = Chunk::new("raw_i32_function_branch");
    c.func_sig = Some((String::new(), "i32".into()));
    equal(&mut c);
    c.emit_br(0, 1);
    validate(c);
}

#[test]
fn i32_loop_result_does_not_type_its_parameterless_branch_label() {
    let mut c = Chunk::new("i32_loop_label");
    c.local_count = 2;
    c.emit_i32_const(0, 1);
    c.emit_op_u16(Op::LOCAL_SET, 0, 1);
    let object = c.add_import("ecma:object", "new");
    c.emit_call(object, 0, 1);
    c.emit_op_u16(Op::LOCAL_SET, 1, 1);
    let loop_offset = c.code.len();
    c.emit_loop_typed(1, 1);
    c.block_i32_results.insert(loop_offset);
    c.emit_op_u16(Op::LOCAL_GET, 0, 1);
    c.emit_op(Op::I32_EQZ, 1);
    c.emit_if(1);
    c.emit_i32_const(1, 1);
    c.emit_op_u16(Op::LOCAL_SET, 0, 1);
    // The loop label consumes no payload. This discarded object must never
    // be unboxed as i32 just because the loop's eventual result is i32.
    c.emit_op_u16(Op::LOCAL_GET, 1, 1);
    c.emit_br(1, 1);
    c.emit_end(1);
    c.emit_i32_const(7, 1);
    c.emit_end(1);
    c.emit_op(Op::RETURN, 1);
    let bytes = vybe_platform_wasm::write_wasm(&[c]);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    export("i32_loop_label", &bytes);
}
