//! Blocktypes that take PARAMETERS.
//!
//! `block`/`loop`/`if`/`try_table` may take operands as well as produce them
//! (spec §2.4.7, blocktype `bt`). The label's operand-stack base then sits
//! BELOW those params: a `br` keeps the label's arity off the top and discards
//! everything down to that base, params included.
//!
//! Nothing in `tests/wast` reaches this. The `.wast` front end lowers block
//! results through scratch locals and emits every blocktype as `(0 params, N
//! results)`, so a param'd block never gets encoded there — a corpus-wide
//! disassembly finds no nonzero param byte anywhere. The only producer of a
//! nonzero param count is `reader/mod.rs`, i.e. a module read back from a
//! conforming `.wasm`. These tests drive that encoding directly.
//!
//! Expected values are wasmtime's, taken from the equivalent `.wat`:
//!
//! ```wat
//! (func (export "blockparam") (result i32)
//!   (i32.const 1000)
//!   (i32.const 7)
//!   (block (param i32) (result i32)
//!     (i32.const 111)
//!     (i32.const 42)
//!     (br 0))
//!   (i32.add))                            ;; => 1042
//! ```

use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};

fn run_chunk_expect_i32(chunk: Chunk, expected: i32, what: &str) {
    match VM::new().run(vec![chunk]).expect("chunk should execute") {
        Value::I32(n) => assert_eq!(n, expected, "{what}"),
        other => panic!("{what}: expected i32 {expected}, got {other:?}"),
    }
}

/// `br` out of a `block (param i32) (result i32)` must discard the param along
/// with the block-local junk. Recording the label base ABOVE the param leaves
/// the 7 stranded under the result, and the trailing `i32.add` then computes
/// `7 + 42` instead of `1000 + 42`.
#[test]
fn br_out_of_a_param_block_discards_the_param() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_i32_const(1000, 0); // below the block's base — must survive
    chunk.emit_i32_const(7, 0); // the block's PARAM — must be discarded
    chunk.emit_block_params(0, 1, 1);
    chunk.emit_i32_const(111, 0); // block-local junk — must be discarded
    chunk.emit_i32_const(42, 0); // the block's RESULT — must survive
    chunk.emit_br(0, 0);
    chunk.emit_end(0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::RETURN, 0);

    run_chunk_expect_i32(chunk, 1042, "1000 + 42, with the param and junk gone");
}

/// A `br` to a `loop (param i32)` carries the params and restarts at the same
/// base. If the base is recorded above the params, every iteration re-pushes
/// them and the operand stack grows without bound.
///
/// The loop counts 3 → 0, leaving junk below the carried param on each pass.
/// A base one slot too high keeps one junk value per iteration, so the final
/// `i32.add` reads a leaked 111 instead of the 1000.
#[test]
fn br_to_a_param_loop_does_not_grow_the_stack() {
    // Mirrors this, which wasmtime runs to 1000:
    //
    //   (func (export "loopparam") (param $n i32) (result i32)
    //     (i32.const 1000)
    //     (block $done (result i32)
    //       (local.get $n)
    //       (loop $l (param i32) (result i32)
    //         (i32.const 1) (i32.sub) (local.tee $n)
    //         (br_if $done (i32.le_s (local.get $n) (i32.const 0)))
    //         (i32.const 111)                 ;; junk, below the carried param
    //         (local.get $n)
    //         (br $l)))
    //     (i32.add))
    let mut chunk = Chunk::new("<script>");
    // A script frame has base 0, so local slot 0 IS operand-stack slot 0.
    // Reserve it for `$n` before pushing anything that has to survive,
    // otherwise `local.tee 0` overwrites the value below the block.
    chunk.emit_i32_const(0, 0); // slot 0 == local $n
    chunk.emit_i32_const(1000, 0); // below everything — must survive
    chunk.emit_block_params(0, 0, 1); // $done (result i32)
    chunk.emit_i32_const(3, 0); // the loop's initial param
    chunk.emit_loop_params(0, 1, 1);

    // stack: [n]
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::I32_SUB, 0); // [n-1]
    chunk.emit_op_u16(Op::LOCAL_TEE, 0, 0); // $n = n-1, stack [n-1]
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op(Op::I32_LE_S, 0); // [n-1, cond]
    chunk.emit_br_if(1, 0); // exit to $done carrying n-1

    chunk.emit_i32_const(111, 0); // junk — `br $l` must discard it
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0); // the value carried as the param
    chunk.emit_br(0, 0);
    chunk.emit_end(0); // end loop
    chunk.emit_end(0); // end block
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::RETURN, 0);

    run_chunk_expect_i32(chunk, 1000, "1000 + 0, with no junk leaked per iteration");
}
