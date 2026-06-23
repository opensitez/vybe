use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

const RATE_EPSILON: f64 = 1e-12;
const SOLVER_EPSILON: f64 = 1e-10;
const DERIVATIVE_EPSILON: f64 = 1e-7;
const MAX_RATE_ITERATIONS: f64 = 25.0;

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let slot = chunk.local_count;
    chunk.local_count = slot + 1;
    slot
}

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

fn push_f64(chunk: &mut Chunk, value: f64, line: u32) {
    push_const(chunk, Value::F64(value), line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_op(Op::DROP, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[0].add_import(module.to_string(), name.to_string());
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(argc, line);
}

fn emit_log(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "ecma:math", "log", 1, line);
}

fn emit_pow_one_plus_local(
    chunks: &mut [Chunk],
    current: usize,
    rate_slot: u16,
    exp_slot: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    lget(chunk, rate_slot, line);
    push_f64(chunk, 1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    lget(chunk, exp_slot, line);
    call_import(chunks, current, "ecma:math", "pow", 2, line);
}

fn emit_local_abs_lt(chunk: &mut Chunk, slot: u16, limit: f64, line: u32) {
    lget(chunk, slot, line);
    chunk.emit_op(Op::F64_ABS, line);
    push_f64(chunk, limit, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
}

fn init_slots_3(chunk: &mut Chunk, line: u32) -> [u16; 3] {
    let s1 = alloc_local(chunk);
    let s2 = alloc_local(chunk);
    let s3 = alloc_local(chunk);
    lset(chunk, s3, line);
    lset(chunk, s2, line);
    lset(chunk, s1, line);
    [s1, s2, s3]
}

fn init_slots_4(chunk: &mut Chunk, line: u32) -> [u16; 4] {
    let s1 = alloc_local(chunk);
    let s2 = alloc_local(chunk);
    let s3 = alloc_local(chunk);
    let s4 = alloc_local(chunk);
    lset(chunk, s4, line);
    lset(chunk, s3, line);
    lset(chunk, s2, line);
    lset(chunk, s1, line);
    [s1, s2, s3, s4]
}

fn init_slots_5(chunk: &mut Chunk, argc: u8, default4: f64, default5: f64, line: u32) -> [u16; 5] {
    let s1 = alloc_local(chunk);
    let s2 = alloc_local(chunk);
    let s3 = alloc_local(chunk);
    let s4 = alloc_local(chunk);
    let s5 = alloc_local(chunk);
    match argc {
        5 => {
            lset(chunk, s5, line);
            lset(chunk, s4, line);
            lset(chunk, s3, line);
            lset(chunk, s2, line);
            lset(chunk, s1, line);
        }
        4 => {
            lset(chunk, s4, line);
            lset(chunk, s3, line);
            lset(chunk, s2, line);
            lset(chunk, s1, line);
            push_f64(chunk, default5, line);
            lset(chunk, s5, line);
        }
        _ => {
            lset(chunk, s3, line);
            lset(chunk, s2, line);
            lset(chunk, s1, line);
            push_f64(chunk, default4, line);
            lset(chunk, s4, line);
            push_f64(chunk, default5, line);
            lset(chunk, s5, line);
        }
    }
    [s1, s2, s3, s4, s5]
}

fn init_slots_5_required4(chunk: &mut Chunk, argc: u8, default5: f64, line: u32) -> [u16; 5] {
    let s1 = alloc_local(chunk);
    let s2 = alloc_local(chunk);
    let s3 = alloc_local(chunk);
    let s4 = alloc_local(chunk);
    let s5 = alloc_local(chunk);
    match argc {
        5 => {
            lset(chunk, s5, line);
            lset(chunk, s4, line);
            lset(chunk, s3, line);
            lset(chunk, s2, line);
            lset(chunk, s1, line);
        }
        _ => {
            lset(chunk, s4, line);
            lset(chunk, s3, line);
            lset(chunk, s2, line);
            lset(chunk, s1, line);
            push_f64(chunk, default5, line);
            lset(chunk, s5, line);
        }
    }
    [s1, s2, s3, s4, s5]
}

fn init_slots_6_required3(
    chunk: &mut Chunk,
    argc: u8,
    default4: f64,
    default5: f64,
    default6: f64,
    line: u32,
) -> [u16; 6] {
    let s1 = alloc_local(chunk);
    let s2 = alloc_local(chunk);
    let s3 = alloc_local(chunk);
    let s4 = alloc_local(chunk);
    let s5 = alloc_local(chunk);
    let s6 = alloc_local(chunk);
    match argc {
        6 => {
            lset(chunk, s6, line);
            lset(chunk, s5, line);
            lset(chunk, s4, line);
            lset(chunk, s3, line);
            lset(chunk, s2, line);
            lset(chunk, s1, line);
        }
        5 => {
            lset(chunk, s5, line);
            lset(chunk, s4, line);
            lset(chunk, s3, line);
            lset(chunk, s2, line);
            lset(chunk, s1, line);
            push_f64(chunk, default6, line);
            lset(chunk, s6, line);
        }
        4 => {
            lset(chunk, s4, line);
            lset(chunk, s3, line);
            lset(chunk, s2, line);
            lset(chunk, s1, line);
            push_f64(chunk, default5, line);
            lset(chunk, s5, line);
            push_f64(chunk, default6, line);
            lset(chunk, s6, line);
        }
        _ => {
            lset(chunk, s3, line);
            lset(chunk, s2, line);
            lset(chunk, s1, line);
            push_f64(chunk, default4, line);
            lset(chunk, s4, line);
            push_f64(chunk, default5, line);
            lset(chunk, s5, line);
            push_f64(chunk, default6, line);
            lset(chunk, s6, line);
        }
    }
    [s1, s2, s3, s4, s5, s6]
}

fn init_slots_6_required4(
    chunk: &mut Chunk,
    argc: u8,
    default5: f64,
    default6: f64,
    line: u32,
) -> [u16; 6] {
    let s1 = alloc_local(chunk);
    let s2 = alloc_local(chunk);
    let s3 = alloc_local(chunk);
    let s4 = alloc_local(chunk);
    let s5 = alloc_local(chunk);
    let s6 = alloc_local(chunk);
    match argc {
        6 => {
            lset(chunk, s6, line);
            lset(chunk, s5, line);
            lset(chunk, s4, line);
            lset(chunk, s3, line);
            lset(chunk, s2, line);
            lset(chunk, s1, line);
        }
        5 => {
            lset(chunk, s5, line);
            lset(chunk, s4, line);
            lset(chunk, s3, line);
            lset(chunk, s2, line);
            lset(chunk, s1, line);
            push_f64(chunk, default6, line);
            lset(chunk, s6, line);
        }
        _ => {
            lset(chunk, s4, line);
            lset(chunk, s3, line);
            lset(chunk, s2, line);
            lset(chunk, s1, line);
            push_f64(chunk, default5, line);
            lset(chunk, s5, line);
            push_f64(chunk, default6, line);
            lset(chunk, s6, line);
        }
    }
    [s1, s2, s3, s4, s5, s6]
}

fn emit_pmt_formula(
    chunks: &mut [Chunk],
    current: usize,
    rate: u16,
    nper: u16,
    pv: u16,
    fv: u16,
    typ: u16,
    line: u32,
) {
    {
        let chunk = &mut chunks[current];
        emit_local_abs_lt(chunk, rate, RATE_EPSILON, line);
        chunk.emit_if(line);
    }
    {
        let chunk = &mut chunks[current];
        lget(chunk, pv, line);
        lget(chunk, fv, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_op(Op::F64_NEG, line);
        lget(chunk, nper, line);
        chunk.emit_op(Op::F64_DIV, line);
        chunk.emit_else(line);
    }

    let pow_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    emit_pow_one_plus_local(chunks, current, rate, nper, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, pow_slot, line);
        lget(chunk, fv, line);
        lget(chunk, pv, line);
        lget(chunk, pow_slot, line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_op(Op::F64_NEG, line);
        lget(chunk, rate, line);
        chunk.emit_op(Op::F64_MUL, line);
        lget(chunk, rate, line);
        lget(chunk, typ, line);
        chunk.emit_op(Op::F64_MUL, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_ADD, line);
        lget(chunk, pow_slot, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_SUB, line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_DIV, line);
        chunk.emit_end(line);
    }
}

fn emit_rate_equation_value(
    chunks: &mut [Chunk],
    current: usize,
    rate: u16,
    nper: u16,
    pmt: u16,
    pv: u16,
    fv: u16,
    typ: u16,
    line: u32,
) {
    {
        let chunk = &mut chunks[current];
        emit_local_abs_lt(chunk, rate, RATE_EPSILON, line);
        chunk.emit_if(line);
    }
    {
        let chunk = &mut chunks[current];
        lget(chunk, fv, line);
        lget(chunk, pv, line);
        chunk.emit_op(Op::F64_ADD, line);
        lget(chunk, pmt, line);
        lget(chunk, nper, line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_else(line);
    }

    let pow_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    emit_pow_one_plus_local(chunks, current, rate, nper, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, pow_slot, line);
        lget(chunk, fv, line);
        lget(chunk, pv, line);
        lget(chunk, pow_slot, line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_ADD, line);
        lget(chunk, pmt, line);
        lget(chunk, rate, line);
        lget(chunk, typ, line);
        chunk.emit_op(Op::F64_MUL, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_op(Op::F64_MUL, line);
        lget(chunk, pow_slot, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_SUB, line);
        chunk.emit_op(Op::F64_MUL, line);
        lget(chunk, rate, line);
        chunk.emit_op(Op::F64_DIV, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_end(line);
    }
}

fn emit_ipmt_formula(
    chunks: &mut [Chunk],
    current: usize,
    rate: u16,
    per: u16,
    pv: u16,
    typ: u16,
    payment_slot: u16,
    line: u32,
) {
    {
        let chunk = &mut chunks[current];
        emit_local_abs_lt(chunk, rate, RATE_EPSILON, line);
        chunk.emit_if(line);
        push_f64(chunk, 0.0, line);
        chunk.emit_else(line);
    }

    let exp_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    let pow_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };

    {
        let chunk = &mut chunks[current];
        lget(chunk, typ, line);
        push_f64(chunk, 0.0, line);
        crate::emitter::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if(line);
    }

    {
        let chunk = &mut chunks[current];
        lget(chunk, per, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_SUB, line);
        lset(chunk, exp_slot, line);
    }
    emit_pow_one_plus_local(chunks, current, rate, exp_slot, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, pow_slot, line);
        lget(chunk, pv, line);
        lget(chunk, pow_slot, line);
        chunk.emit_op(Op::F64_MUL, line);
        lget(chunk, payment_slot, line);
        lget(chunk, pow_slot, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_SUB, line);
        chunk.emit_op(Op::F64_MUL, line);
        lget(chunk, rate, line);
        chunk.emit_op(Op::F64_DIV, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_op(Op::F64_NEG, line);
        lget(chunk, rate, line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_else(line);
    }

    {
        let chunk = &mut chunks[current];
        lget(chunk, per, line);
        push_f64(chunk, 1.0, line);
        crate::emitter::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if(line);
        push_f64(chunk, 0.0, line);
        chunk.emit_else(line);
    }

    {
        let chunk = &mut chunks[current];
        lget(chunk, per, line);
        push_f64(chunk, 2.0, line);
        chunk.emit_op(Op::F64_SUB, line);
        lset(chunk, exp_slot, line);
    }
    emit_pow_one_plus_local(chunks, current, rate, exp_slot, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, pow_slot, line);
        lget(chunk, pv, line);
        lget(chunk, payment_slot, line);
        chunk.emit_op(Op::F64_ADD, line);
        lget(chunk, pow_slot, line);
        chunk.emit_op(Op::F64_MUL, line);
        lget(chunk, payment_slot, line);
        lget(chunk, pow_slot, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_SUB, line);
        chunk.emit_op(Op::F64_MUL, line);
        lget(chunk, rate, line);
        chunk.emit_op(Op::F64_DIV, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_op(Op::F64_NEG, line);
        lget(chunk, rate, line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_end(line);
        chunk.emit_end(line);
    }
}

pub fn emit_vb_pmt(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let [rate, nper, pv, fv, typ] = {
        let chunk = &mut chunks[current];
        init_slots_5(chunk, argc, 0.0, 0.0, line)
    };
    emit_pmt_formula(chunks, current, rate, nper, pv, fv, typ, line);
}

pub fn emit_vb_fv(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let [rate, nper, pmt, pv, typ] = {
        let chunk = &mut chunks[current];
        init_slots_5(chunk, argc, 0.0, 0.0, line)
    };
    {
        let chunk = &mut chunks[current];
        emit_local_abs_lt(chunk, rate, RATE_EPSILON, line);
        chunk.emit_if(line);
    }
    {
        let chunk = &mut chunks[current];
        lget(chunk, pv, line);
        lget(chunk, pmt, line);
        lget(chunk, nper, line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_op(Op::F64_NEG, line);
        chunk.emit_else(line);
    }

    let pow_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    emit_pow_one_plus_local(chunks, current, rate, nper, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, pow_slot, line);
        lget(chunk, pv, line);
        lget(chunk, pow_slot, line);
        chunk.emit_op(Op::F64_MUL, line);
        lget(chunk, pmt, line);
        lget(chunk, rate, line);
        lget(chunk, typ, line);
        chunk.emit_op(Op::F64_MUL, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_op(Op::F64_MUL, line);
        lget(chunk, pow_slot, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_SUB, line);
        chunk.emit_op(Op::F64_MUL, line);
        lget(chunk, rate, line);
        chunk.emit_op(Op::F64_DIV, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_op(Op::F64_NEG, line);
        chunk.emit_end(line);
    }
}

pub fn emit_vb_pv(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let [rate, nper, pmt, fv, typ] = {
        let chunk = &mut chunks[current];
        init_slots_5(chunk, argc, 0.0, 0.0, line)
    };
    {
        let chunk = &mut chunks[current];
        emit_local_abs_lt(chunk, rate, RATE_EPSILON, line);
        chunk.emit_if(line);
    }
    {
        let chunk = &mut chunks[current];
        lget(chunk, fv, line);
        lget(chunk, pmt, line);
        lget(chunk, nper, line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_op(Op::F64_NEG, line);
        chunk.emit_else(line);
    }

    let pow_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    emit_pow_one_plus_local(chunks, current, rate, nper, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, pow_slot, line);
        lget(chunk, fv, line);
        lget(chunk, pmt, line);
        lget(chunk, rate, line);
        lget(chunk, typ, line);
        chunk.emit_op(Op::F64_MUL, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_op(Op::F64_MUL, line);
        lget(chunk, pow_slot, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_SUB, line);
        chunk.emit_op(Op::F64_MUL, line);
        lget(chunk, rate, line);
        chunk.emit_op(Op::F64_DIV, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_op(Op::F64_NEG, line);
        lget(chunk, pow_slot, line);
        chunk.emit_op(Op::F64_DIV, line);
        chunk.emit_end(line);
    }
}

pub fn emit_vb_nper(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let [rate, pmt, pv, fv, typ] = {
        let chunk = &mut chunks[current];
        init_slots_5(chunk, argc, 0.0, 0.0, line)
    };
    {
        let chunk = &mut chunks[current];
        emit_local_abs_lt(chunk, rate, RATE_EPSILON, line);
        chunk.emit_if(line);
    }
    {
        let chunk = &mut chunks[current];
        lget(chunk, fv, line);
        lget(chunk, pv, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_op(Op::F64_NEG, line);
        lget(chunk, pmt, line);
        chunk.emit_op(Op::F64_DIV, line);
        chunk.emit_else(line);
    }

    let factor_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    {
        let chunk = &mut chunks[current];
        lget(chunk, rate, line);
        lget(chunk, typ, line);
        chunk.emit_op(Op::F64_MUL, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, factor_slot, line);

        lget(chunk, pmt, line);
        lget(chunk, factor_slot, line);
        chunk.emit_op(Op::F64_MUL, line);
        lget(chunk, fv, line);
        lget(chunk, rate, line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_SUB, line);

        lget(chunk, pv, line);
        lget(chunk, rate, line);
        chunk.emit_op(Op::F64_MUL, line);
        lget(chunk, pmt, line);
        lget(chunk, factor_slot, line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_op(Op::F64_DIV, line);
    }
    emit_log(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, rate, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_ADD, line);
    }
    emit_log(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::F64_DIV, line);
        chunk.emit_end(line);
    }
}

pub fn emit_vb_rate(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let [nper, pmt, pv, fv, typ, guess] = {
        let chunk = &mut chunks[current];
        init_slots_6_required3(chunk, argc, 0.0, 0.0, 0.1, line)
    };

    let rate_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    let plus_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    let f_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    let f2_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    let deriv_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    let iter_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };

    {
        let chunk = &mut chunks[current];
        lget(chunk, guess, line);
        lset(chunk, rate_slot, line);
        push_f64(chunk, 0.0, line);
        lset(chunk, iter_slot, line);
    }

    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, iter_slot, line);
        push_f64(chunk, MAX_RATE_ITERATIONS, line);
        crate::emitter::ops::emit_dyn_lt(chunk, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_br_if(1, line);
    }

    emit_rate_equation_value(chunks, current, rate_slot, nper, pmt, pv, fv, typ, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, f_slot, line);
        lget(chunk, f_slot, line);
        chunk.emit_op(Op::F64_ABS, line);
        push_f64(chunk, SOLVER_EPSILON, line);
        crate::emitter::ops::emit_dyn_lt(chunk, line);
        chunk.emit_br_if(1, line);
    }

    {
        let chunk = &mut chunks[current];
        lget(chunk, rate_slot, line);
        push_f64(chunk, DERIVATIVE_EPSILON, line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, plus_slot, line);
    }
    emit_rate_equation_value(chunks, current, plus_slot, nper, pmt, pv, fv, typ, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, f2_slot, line);
        lget(chunk, f2_slot, line);
        lget(chunk, f_slot, line);
        chunk.emit_op(Op::F64_SUB, line);
        push_f64(chunk, DERIVATIVE_EPSILON, line);
        chunk.emit_op(Op::F64_DIV, line);
        lset(chunk, deriv_slot, line);
        lget(chunk, deriv_slot, line);
        chunk.emit_op(Op::F64_ABS, line);
        push_f64(chunk, RATE_EPSILON, line);
        crate::emitter::ops::emit_dyn_lt(chunk, line);
        chunk.emit_br_if(1, line);
    }

    {
        let chunk = &mut chunks[current];
        lget(chunk, rate_slot, line);
        lget(chunk, f_slot, line);
        lget(chunk, deriv_slot, line);
        chunk.emit_op(Op::F64_DIV, line);
        chunk.emit_op(Op::F64_SUB, line);
        lset(chunk, rate_slot, line);
        lget(chunk, iter_slot, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, iter_slot, line);
        chunk.emit_br(0, line);
        chunk.emit_end(line);
        chunk.patch_loop(loop_patch);
    }
    lget(&mut chunks[current], rate_slot, line);
}

pub fn emit_vb_ipmt(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let [rate, per, nper, pv, fv, typ] = {
        let chunk = &mut chunks[current];
        init_slots_6_required4(chunk, argc, 0.0, 0.0, line)
    };
    let payment_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    emit_pmt_formula(chunks, current, rate, nper, pv, fv, typ, line);
    lset(&mut chunks[current], payment_slot, line);
    emit_ipmt_formula(chunks, current, rate, per, pv, typ, payment_slot, line);
}

pub fn emit_vb_ppmt(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let [rate, per, nper, pv, fv, typ] = {
        let chunk = &mut chunks[current];
        init_slots_6_required4(chunk, argc, 0.0, 0.0, line)
    };
    let payment_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    let ipmt_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    emit_pmt_formula(chunks, current, rate, nper, pv, fv, typ, line);
    lset(&mut chunks[current], payment_slot, line);
    emit_ipmt_formula(chunks, current, rate, per, pv, typ, payment_slot, line);
    lset(&mut chunks[current], ipmt_slot, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, payment_slot, line);
        lget(chunk, ipmt_slot, line);
        chunk.emit_op(Op::F64_SUB, line);
    }
}

pub fn emit_vb_sln(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let [cost, salvage, life] = {
        let chunk = &mut chunks[current];
        init_slots_3(chunk, line)
    };
    {
        let chunk = &mut chunks[current];
        lget(chunk, cost, line);
        lget(chunk, salvage, line);
        chunk.emit_op(Op::F64_SUB, line);
        lget(chunk, life, line);
        chunk.emit_op(Op::F64_DIV, line);
    }
}

pub fn emit_vb_ddb(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let [cost, salvage, life, period, factor] = {
        let chunk = &mut chunks[current];
        init_slots_5_required4(chunk, argc, 2.0, line)
    };
    let rate_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    let pow_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    let book_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    {
        let chunk = &mut chunks[current];
        lget(chunk, factor, line);
        lget(chunk, life, line);
        chunk.emit_op(Op::F64_DIV, line);
        lset(chunk, rate_slot, line);
        push_f64(chunk, 1.0, line);
        lget(chunk, rate_slot, line);
        chunk.emit_op(Op::F64_SUB, line);
        lget(chunk, period, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_SUB, line);
    }
    call_import(chunks, current, "ecma:math", "pow", 2, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, pow_slot, line);
        lget(chunk, cost, line);
        lget(chunk, pow_slot, line);
        chunk.emit_op(Op::F64_MUL, line);
        lset(chunk, book_slot, line);
        lget(chunk, book_slot, line);
        lget(chunk, rate_slot, line);
        chunk.emit_op(Op::F64_MUL, line);
        lget(chunk, book_slot, line);
        lget(chunk, salvage, line);
        chunk.emit_op(Op::F64_SUB, line);
        chunk.emit_op(Op::F64_MIN, line);
        push_f64(chunk, 0.0, line);
        chunk.emit_op(Op::F64_MAX, line);
    }
}

pub fn emit_vb_syd(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let [cost, salvage, life, period] = {
        let chunk = &mut chunks[current];
        init_slots_4(chunk, line)
    };
    {
        let chunk = &mut chunks[current];
        lget(chunk, cost, line);
        lget(chunk, salvage, line);
        chunk.emit_op(Op::F64_SUB, line);
        lget(chunk, life, line);
        lget(chunk, period, line);
        chunk.emit_op(Op::F64_SUB, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_op(Op::F64_MUL, line);
        push_f64(chunk, 2.0, line);
        chunk.emit_op(Op::F64_MUL, line);
        lget(chunk, life, line);
        lget(chunk, life, line);
        push_f64(chunk, 1.0, line);
        chunk.emit_op(Op::F64_ADD, line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_DIV, line);
    }
}
