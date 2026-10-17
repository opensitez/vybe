#![cfg(feature = "primitive-benchmarks")]
use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op};
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;

#[derive(Clone)]
enum Expr {
    Get(u16),
    Const(i32),
    Unary(u16, Box<Expr>),
    Binary(u16, Box<Expr>, Box<Expr>),
    Store(Box<Expr>),
    Call(Box<Expr>),
}
fn emit(c: &mut Chunk, expr: &Expr, barriers: bool, tokens: &mut Vec<serde_json::Value>) {
    match expr {
        Expr::Get(slot) => {
            c.emit_op_u16(Op::LOCAL_GET, *slot, 0);
            tokens.push(serde_json::json!(["get", slot]));
        }
        Expr::Const(value) => {
            if barriers {
                c.emit_op(Op::NOP, 0);
            }
            c.emit_i32_const(*value, 0);
            tokens.push(serde_json::json!(["const", value]));
        }
        Expr::Unary(op, a) => {
            emit(c, a, barriers, tokens);
            if barriers {
                c.emit_op(Op::NOP, 0);
            }
            c.emit_op(Op::decode(0, *op).unwrap(), 0);
            tokens.push(serde_json::json!(["op", op]));
        }
        Expr::Binary(op, a, b) => {
            emit(c, a, barriers, tokens);
            emit(c, b, barriers, tokens);
            if barriers {
                c.emit_op(Op::NOP, 0);
            }
            c.emit_op(Op::decode(0, *op).unwrap(), 0);
            tokens.push(serde_json::json!(["op", op]));
        }
        Expr::Store(a) => {
            emit(c, a, barriers, tokens);
            c.emit_op_u16(Op::LOCAL_SET, 3, 0);
            c.emit_op_u16(Op::LOCAL_GET, 3, 0);
            tokens.push(serde_json::json!(["set", 3]));
            tokens.push(serde_json::json!(["get", 3]));
        }
        Expr::Call(a) => {
            emit(c, a, barriers, tokens);
            let import = c.add_import("wasm:js-number", "fromI32");
            c.emit_call(import, 1, 0);
            tokens.push(serde_json::json!(["call", 0]));
        }
    }
}
fn binary(op: u16, a: Expr, b: Expr) -> Expr {
    Expr::Binary(op, Box::new(a), Box::new(b))
}
fn random(seed: &mut u32, depth: u32) -> Expr {
    *seed = seed.wrapping_mul(1664525).wrapping_add(1013904223);
    let choice = *seed;
    if depth == 0 {
        return if choice & 1 == 0 {
            Expr::Get((choice % 3) as u16)
        } else {
            Expr::Const([0, 1, -1, 31, i32::MIN, i32::MAX][(choice as usize >> 2) % 6])
        };
    }
    if choice % 7 == 0 {
        return Expr::Store(Box::new(random(seed, depth - 1)));
    }
    if choice % 11 == 0 {
        return Expr::Call(Box::new(random(seed, depth - 1)));
    }
    if choice % 5 == 0 {
        return Expr::Unary(
            [0x45, 0x67, 0x68, 0x69][(choice as usize >> 3) % 4],
            Box::new(random(seed, depth - 1)),
        );
    }
    let op = if choice & 1 == 0 {
        0x46 + (choice % 10) as u16
    } else {
        0x6a + (choice % 15) as u16
    };
    binary(op, random(seed, depth - 1), random(seed, depth - 1))
}
#[test]
fn nested_expressions_and_abi_boundaries_export_valid_modules() {
    let mut expressions = vec![
        binary(0x6a, Expr::Get(0), binary(0x6d, Expr::Get(1), Expr::Get(2))),
        binary(0x6a, binary(0x6d, Expr::Get(0), Expr::Get(1)), Expr::Get(2)),
        binary(
            0x6a,
            binary(0x6a, Expr::Get(0), Expr::Const(1)),
            binary(0x4f, Expr::Get(1), Expr::Const(128)),
        ),
    ];
    let mut seed = 0x8c170b35;
    for _ in 0..96 {
        expressions.push(random(&mut seed, 3));
    }
    let mut rows = vec![];
    for (index, expr) in expressions.iter().enumerate() {
        for condition in [false, true] {
            let mut names = vec![];
            let mut reference = vec![];
            for barriers in [true, false] {
                let mut root = Chunk::new("owner");
                primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
                root.emit_op(Op::RETURN, 0);
                let mut entry = Chunk::new("entry");
                entry.arity = 3;
                entry.param_count = 3;
                entry.local_count = 4;
                let mut tokens = vec![];
                emit(&mut entry, expr, barriers, &mut tokens);
                if condition {
                    // Calls can return boxed numeric values. Establish the
                    // i32 condition transport before the typed Wasm `if`.
                    if barriers { entry.emit_op(Op::NOP,0); }
                    entry.emit_op(Op::I32_EQZ,0);
                    if barriers { entry.emit_op(Op::NOP,0); }
                    entry.emit_op(Op::I32_EQZ,0);
                    entry.emit_if_i32(0);
                    entry.emit_i32_const(1, 0);
                    entry.emit_else(0);
                    entry.emit_i32_const(0, 0);
                    entry.emit_end(0);
                }
                entry.emit_op(Op::RETURN, 0);
                let name = format!("region-{index}-{condition}-{barriers}");
                native_export::export_chunks(&[root, entry], &name);
                names.push(name);
                reference = tokens;
            }
            rows.push(serde_json::json!({"names":names,"tokens":reference,"condition":condition}));
        }
    }
    std::fs::write(
        std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap())
            .join("I32_REGION_CASES.json"),
        serde_json::to_vec_pretty(&rows).unwrap(),
    )
    .unwrap();
}
