#![cfg(feature = "primitive-benchmarks")]
use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op};
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;

#[test]
fn every_literal_binary_operator_has_a_valid_complete_module() {
    let mut rows = vec![];
    for opcode in (0x46u16..=0x4f).chain(0x6a..=0x78) {
        for rhs in [0, 1, -1, 31, i32::MIN, i32::MAX] {
            for condition in [false, true] {
                let mut root = Chunk::new("owner");
                primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
                root.emit_op(Op::RETURN, 0);
                let mut entry = Chunk::new("entry");
                entry.arity = 1;
                entry.param_count = 1;
                entry.local_count = 1;
                entry.emit_op_u16(Op::LOCAL_GET, 0, 0);
                entry.emit_i32_const(rhs, 0);
                entry.emit_op(Op::decode(0, opcode).unwrap(), 0);
                if condition {
                    entry.emit_if_i32(0);
                    entry.emit_i32_const(1, 0);
                    entry.emit_else(0);
                    entry.emit_i32_const(0, 0);
                    entry.emit_end(0);
                }
                entry.emit_op(Op::RETURN, 0);
                let name = format!("literal-{opcode}-{rhs}-{condition}");
                native_export::export_chunks(&[root, entry], &name);
                rows.push(serde_json::json!({"name":name,"opcode":opcode,"rhs":rhs,"condition":condition}));
            }
        }
    }
    std::fs::write(
        std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap())
            .join("LITERAL_RHS_CASES.json"),
        serde_json::to_vec_pretty(&rows).unwrap(),
    )
    .unwrap();
}
