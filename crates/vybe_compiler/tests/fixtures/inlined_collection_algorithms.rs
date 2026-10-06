//! Fixed pre-outlining emission reference for module-size comparisons.
//! Algorithms are identical; only their packaging at call sites differs.
use vybe_compiler::primitives::collections::{
    emit_get_into, emit_len_into, emit_new_with_length_into, emit_set_into,
};
use vybe_compiler::primitives::{collections, ops};
use vybe_runtime::{Chunk, Op};
fn lget(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
}
fn lset(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
}
pub fn emit_sort_by_key_in_place(chunks: &mut [Chunk], current: usize, line: u32) {
    let abi = vybe_compiler::primitives::class_context::module_receiver_abi(chunks);
    let c = &mut chunks[current];
    let base = c.alloc_scratch(6);
    let key_fn = base;
    let arr = base + 1;
    let len = base + 2;
    let index = base + 3;
    let decorated = base + 4;
    let value = base + 5;
    lset(c, key_fn, line);
    lset(c, arr, line);
    let mut imports = Chunk::new("imports");
    lget(c, arr, line);
    emit_len_into(&mut imports, c, line);
    lset(c, len, line);
    lget(c, len, line);
    emit_new_with_length_into(&mut imports, c, line);
    lset(c, decorated, line);

    // Evaluate keys exactly once, in input order. Keep the original values in
    // private fixed-size GC pairs, so a throwing key never partially sorts arr.
    c.emit_f64_const(0.0, line);
    lset(c, index, line);
    let done = c.emit_block(line);
    let (repeat, _) = c.emit_loop_s(line);
    lget(c, index, line);
    lget(c, len, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_br_if(1, line);
    lget(c, arr, line);
    lget(c, index, line);
    emit_get_into(&mut imports, c, line);
    lset(c, value, line);
    lget(c, decorated, line);
    lget(c, index, line);
    lget(c, key_fn, line);
    let recv = vybe_compiler::primitives::callable::emit_callback_receiver(c, abi, line);
    lget(c, value, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(c, 1 + recv, line);
    lget(c, value, line);
    c.emit_array_new_fixed(0, 2, line);
    emit_set_into(&mut imports, c, line);
    c.emit_op(Op::DROP, line);
    lget(c, index, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_ADD, line);
    lset(c, index, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(repeat);
    c.emit_end(line);
    c.patch_block(done);

    emit_stable_merge_sort(&mut imports, c, decorated, line, |_, c, lhs, rhs, line| {
        for pair in [lhs, rhs] {
            lget(c, pair, line);
            c.emit_i32_const(0, line);
            c.emit_op(Op::ARRAY_GET, line);
        }
        // Preserve the existing dynamic key comparison and coercion contract.
        vybe_compiler::primitives::ops::emit_dyn_gt(c, line);
    });

    c.emit_f64_const(0.0, line);
    lset(c, index, line);
    let done = c.emit_block(line);
    let (repeat, _) = c.emit_loop_s(line);
    lget(c, index, line);
    lget(c, len, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_br_if(1, line);
    lget(c, arr, line);
    lget(c, index, line);
    lget(c, decorated, line);
    lget(c, index, line);
    emit_get_into(&mut imports, c, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::ARRAY_GET, line);
    emit_set_into(&mut imports, c, line);
    c.emit_op(Op::DROP, line);
    lget(c, index, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_ADD, line);
    lset(c, index, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(repeat);
    c.emit_end(line);
    c.patch_block(done);
    lget(c, arr, line);
}
fn emit_stable_merge_sort(
    imports: &mut Chunk,
    c: &mut Chunk,
    arr: u16,
    line: u32,
    mut greater: impl FnMut(&mut Chunk, &mut Chunk, u16, u16, u32),
) {
    let base = c.alloc_scratch(13);
    let len = base;
    let src = base + 1;
    let dst = base + 2;
    let swap = base + 3;
    let width = base + 4;
    let start = base + 5;
    let mid = base + 6;
    let end = base + 7;
    let left = base + 8;
    let right = base + 9;
    let out = base + 10;
    let lhs = base + 11;
    let rhs = base + 12;

    fn increment(c: &mut Chunk, slot: u16, line: u32) {
        lget(c, slot, line);
        c.emit_f64_const(1.0, line);
        c.emit_op(Op::F64_ADD, line);
        lset(c, slot, line);
    }
    fn read(imports: &mut Chunk, c: &mut Chunk, src: u16, index: u16, value: u16, line: u32) {
        lget(c, src, line);
        lget(c, index, line);
        emit_get_into(imports, c, line);
        lset(c, value, line);
    }

    lget(c, arr, line);
    emit_len_into(imports, c, line);
    lset(c, len, line);
    // Empty/singleton arrays need no temporary allocation or comparison.
    let done = c.emit_block(line);
    lget(c, len, line);
    c.emit_f64_const(2.0, line);
    c.emit_op(Op::F64_LT, line);
    c.emit_br_if(0, line);
    // Keep the linear best case of insertion sort. This also avoids a buffer
    // allocation for already sorted arrays, including arrays of equal keys.
    c.emit_f64_const(0.0, line);
    lset(c, left, line);
    read(imports, c, arr, left, lhs, line);
    c.emit_f64_const(1.0, line);
    lset(c, out, line);
    let not_sorted = c.emit_block(line);
    let (scan, _) = c.emit_loop_s(line);
    lget(c, out, line);
    lget(c, len, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_br_if(2, line); // all adjacent pairs ordered: leave done
    read(imports, c, arr, out, rhs, line);
    greater(imports, c, lhs, rhs, line);
    c.emit_br_if(1, line); // first descending pair: merge the array
    lget(c, rhs, line);
    lset(c, lhs, line);
    increment(c, out, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(scan);
    c.emit_end(line);
    c.patch_block(not_sorted);
    lget(c, arr, line);
    lset(c, src, line);
    lget(c, len, line);
    emit_new_with_length_into(imports, c, line);
    lset(c, dst, line);
    c.emit_f64_const(1.0, line);
    lset(c, width, line);

    let passes_done = c.emit_block(line);
    let (passes, _) = c.emit_loop_s(line);
    lget(c, width, line);
    lget(c, len, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_br_if(1, line);
    c.emit_f64_const(0.0, line);
    lset(c, start, line);

    let runs_done = c.emit_block(line);
    let (runs, _) = c.emit_loop_s(line);
    lget(c, start, line);
    lget(c, len, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_br_if(1, line);
    // Bounds remain exact Number integers, including lengths above i32::MAX.
    for (from, to) in [(start, mid), (mid, end)] {
        lget(c, from, line);
        lget(c, width, line);
        c.emit_op(Op::F64_ADD, line);
        lget(c, len, line);
        c.emit_op(Op::F64_MIN, line);
        lset(c, to, line);
    }
    lget(c, start, line);
    lset(c, left, line);
    lget(c, mid, line);
    lset(c, right, line);
    lget(c, start, line);
    lset(c, out, line);

    let merge_done = c.emit_block(line);
    let (merge, _) = c.emit_loop_s(line);
    lget(c, out, line);
    lget(c, end, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_br_if(1, line);
    lget(c, left, line);
    lget(c, mid, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_if(line);
    read(imports, c, src, right, lhs, line);
    increment(c, right, line);
    c.emit_else(line);
    lget(c, right, line);
    lget(c, end, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_if(line);
    read(imports, c, src, left, lhs, line);
    increment(c, left, line);
    c.emit_else(line);
    read(imports, c, src, left, lhs, line);
    read(imports, c, src, right, rhs, line);
    greater(imports, c, lhs, rhs, line);
    c.emit_if(line);
    lget(c, rhs, line);
    lset(c, lhs, line);
    increment(c, right, line);
    c.emit_else(line);
    // Ties select the left run: equal elements keep their original order.
    increment(c, left, line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_end(line);
    lget(c, dst, line);
    lget(c, out, line);
    lget(c, lhs, line);
    emit_set_into(imports, c, line);
    c.emit_op(Op::DROP, line);
    increment(c, out, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(merge);
    c.emit_end(line);
    c.patch_block(merge_done);
    lget(c, end, line);
    lset(c, start, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(runs);
    c.emit_end(line);
    c.patch_block(runs_done);
    lget(c, src, line);
    lset(c, swap, line);
    lget(c, dst, line);
    lset(c, src, line);
    lget(c, swap, line);
    lset(c, dst, line);
    lget(c, width, line);
    c.emit_f64_const(2.0, line);
    c.emit_op(Op::F64_MUL, line);
    lset(c, width, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(passes);
    c.emit_end(line);
    c.patch_block(passes_done);

    lget(c, src, line);
    lget(c, arr, line);
    c.emit_op(Op::REF_EQ, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if(line);
    c.emit_f64_const(0.0, line);
    lset(c, out, line);
    let copy_done = c.emit_block(line);
    let (copy, _) = c.emit_loop_s(line);
    lget(c, out, line);
    lget(c, len, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_br_if(1, line);
    lget(c, arr, line);
    lget(c, out, line);
    lget(c, src, line);
    lget(c, out, line);
    emit_get_into(imports, c, line);
    emit_set_into(imports, c, line);
    c.emit_op(Op::DROP, line);
    increment(c, out, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(copy);
    c.emit_end(line);
    c.patch_block(copy_done);
    c.emit_end(line);
    c.emit_end(line);
    c.patch_block(done);
}
fn get(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
}
fn set(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
}
fn field(c: &mut Chunk, pair: u16, index: i32, line: u32) {
    get(c, pair, line);
    c.emit_i32_const(index, line);
    c.emit_op(Op::ARRAY_GET, line);
}
fn read(c: &mut Chunk, arr: u16, index: u16, value: u16, line: u32) {
    get(c, arr, line);
    get(c, index, line);
    // Dynamic arrays can exceed the signed-i32 range. Keep their numeric
    // indices on the common array surface; GC gets are only used for the
    // three fixed, proven decoration fields.
    let get = c.add_import("ecma:array", "get");
    c.emit_call(get, 2, line);
    set(c, value, line);
}
fn write(c: &mut Chunk, arr: u16, index: u16, value: u16, line: u32) {
    get(c, arr, line);
    get(c, index, line);
    get(c, value, line);
    let set = c.add_import("ecma:array", "set");
    c.emit_call(set, 3, line);
    c.emit_op(Op::DROP, line);
}
fn increment(c: &mut Chunk, index: u16, line: u32) {
    get(c, index, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, index, line);
}
fn key_better(c: &mut Chunk, largest: bool, line: u32) {
    if largest {
        ops::emit_dyn_gt(c, line);
    } else {
        ops::emit_dyn_lt(c, line);
    }
}

// A decoration is [key, original index, value]. The index is a proven
// numeric counter, and makes equal-key selection stable in either direction.
fn pair_better(c: &mut Chunk, lhs: u16, rhs: u16, largest: bool, line: u32) {
    field(c, lhs, 0, line);
    field(c, rhs, 0, line);
    key_better(c, largest, line);
    c.emit_if_i32(line);
    c.emit_i32_const(1, line);
    c.emit_else(line);
    field(c, rhs, 0, line);
    field(c, lhs, 0, line);
    key_better(c, largest, line);
    c.emit_if_i32(line);
    c.emit_i32_const(0, line);
    c.emit_else(line);
    field(c, lhs, 1, line);
    field(c, rhs, 1, line);
    c.emit_op(Op::F64_LT, line);
    c.emit_end(line);
    c.emit_end(line);
}

// Keep the worst retained decoration at the root of the private heap.
fn sift_up(c: &mut Chunk, arr: u16, pos: u16, item: u16, largest: bool, line: u32) {
    let base = c.alloc_scratch(2);
    let parent = base;
    let value = base + 1;
    let done = c.emit_block(line);
    let (repeat, _) = c.emit_loop_s(line);
    get(c, pos, line);
    c.emit_f64_const(0.0, line);
    c.emit_op(Op::F64_LE, line);
    c.emit_br_if(1, line);
    get(c, pos, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_SUB, line);
    c.emit_f64_const(2.0, line);
    c.emit_op(Op::F64_DIV, line);
    c.emit_op(Op::F64_FLOOR, line);
    set(c, parent, line);
    read(c, arr, parent, value, line);
    pair_better(c, value, item, largest, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_br_if(1, line);
    write(c, arr, pos, value, line);
    get(c, parent, line);
    set(c, pos, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(repeat);
    c.emit_end(line);
    c.patch_block(done);
    write(c, arr, pos, item, line);
}

fn sift_down(c: &mut Chunk, arr: u16, len: u16, item: u16, largest: bool, line: u32) {
    let base = c.alloc_scratch(5);
    let pos = base;
    let child = base + 1;
    let right = base + 2;
    let value = base + 3;
    let other = base + 4;
    c.emit_f64_const(0.0, line);
    set(c, pos, line);
    let done = c.emit_block(line);
    let (repeat, _) = c.emit_loop_s(line);
    get(c, pos, line);
    c.emit_f64_const(2.0, line);
    c.emit_op(Op::F64_MUL, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, child, line);
    get(c, child, line);
    get(c, len, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_br_if(1, line);
    read(c, arr, child, value, line);
    get(c, child, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, right, line);
    get(c, right, line);
    get(c, len, line);
    c.emit_op(Op::F64_LT, line);
    c.emit_if(line);
    read(c, arr, right, other, line);
    pair_better(c, value, other, largest, line);
    c.emit_if(line);
    get(c, right, line);
    set(c, child, line);
    get(c, other, line);
    set(c, value, line);
    c.emit_end(line);
    c.emit_end(line);
    pair_better(c, item, value, largest, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_br_if(1, line);
    write(c, arr, pos, value, line);
    get(c, child, line);
    set(c, pos, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(repeat);
    c.emit_end(line);
    c.patch_block(done);
    write(c, arr, pos, item, line);
}

fn emit_n_of(chunks: &mut [Chunk], current: usize, argc: u8, largest: bool, line: u32) {
    let abi = vybe_compiler::primitives::class_context::module_receiver_abi(chunks);
    let c = &mut chunks[current];
    let base = c.alloc_scratch(12);
    let n = base;
    let data = base + 1;
    let key_fn = base + 2;
    let len = base + 3;
    let result = base + 4;
    let index = base + 5;
    let filled = base + 6;
    let pos = base + 7;
    let value = base + 8;
    let key = base + 9;
    let item = base + 10;
    let root = base + 11;
    if argc >= 3 {
        set(c, key_fn, line);
    }
    set(c, data, line);
    set(c, n, line);
    get(c, data, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    set(c, len, line);
    // Normalize the dynamic count once; subsequent operations consume a
    // proven number. Private indices are numeric counters throughout.
    get(c, n, line);
    let to_number = c.add_import("ecma:value", "toNumber");
    c.emit_call(to_number, 1, line);
    set(c, n, line);
    get(c, n, line);
    get(c, n, line);
    c.emit_op(Op::F64_EQ, line);
    c.emit_if_value(line);
    get(c, n, line);
    c.emit_op(Op::F64_FLOOR, line);
    c.emit_else(line);
    c.emit_f64_const(0.0, line);
    c.emit_end(line);
    get(c, len, line);
    c.emit_op(Op::F64_MIN, line);
    c.emit_f64_const(0.0, line);
    c.emit_op(Op::F64_MAX, line);
    set(c, n, line);
    let mut imports = Chunk::new("imports");
    get(c, n, line);
    collections::emit_new_with_length_into(&mut imports, c, line);
    set(c, result, line);
    c.emit_f64_const(0.0, line);
    set(c, index, line);
    c.emit_f64_const(0.0, line);
    set(c, filled, line);
    let done = c.emit_block(line);
    get(c, n, line);
    c.emit_f64_const(0.0, line);
    c.emit_op(Op::F64_LE, line);
    c.emit_br_if(0, line);
    let (repeat, _) = c.emit_loop_s(line);
    get(c, index, line);
    get(c, len, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_br_if(1, line);
    read(c, data, index, value, line);
    if argc >= 3 {
        get(c, key_fn, line);
        c.emit_op(Op::REF_IS_NULL, line);
        c.emit_if_value(line);
        get(c, value, line);
        c.emit_else(line);
        get(c, key_fn, line);
        let recv = vybe_compiler::primitives::callable::emit_callback_receiver(c, abi, line);
        get(c, value, line);
        vybe_compiler::primitives::callable::emit_direct_invoke_chunk(c, 1 + recv, line);
        c.emit_end(line);
    } else {
        get(c, value, line);
    }
    set(c, key, line);
    get(c, filled, line);
    get(c, n, line);
    c.emit_op(Op::F64_LT, line);
    c.emit_if(line);
    get(c, key, line);
    get(c, index, line);
    get(c, value, line);
    c.emit_array_new_fixed(0, 3, line);
    set(c, item, line);
    get(c, filled, line);
    set(c, pos, line);
    // When all elements are requested, decorate directly without maintaining
    // a heap that would immediately be sorted again.
    get(c, n, line);
    get(c, len, line);
    c.emit_op(Op::F64_LT, line);
    c.emit_if(line);
    sift_up(c, result, pos, item, largest, line);
    c.emit_else(line);
    write(c, result, pos, item, line);
    c.emit_end(line);
    increment(c, filled, line);
    c.emit_else(line);
    get(c, result, line);
    c.emit_i32_const(0, line);
    c.emit_op(Op::ARRAY_GET, line);
    set(c, root, line);
    get(c, key, line);
    field(c, root, 0, line);
    key_better(c, largest, line);
    c.emit_if(line);
    get(c, key, line);
    get(c, index, line);
    get(c, value, line);
    c.emit_array_new_fixed(0, 3, line);
    set(c, item, line);
    sift_down(c, result, n, item, largest, line);
    c.emit_end(line);
    c.emit_end(line);
    increment(c, index, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(repeat);
    c.emit_end(line);
    c.patch_block(done);

    emit_stable_merge_sort(&mut imports, c, result, line, |_, c, lhs, rhs, line| {
        pair_better(c, rhs, lhs, largest, line);
    });
    c.emit_f64_const(0.0, line);
    set(c, index, line);
    let done = c.emit_block(line);
    let (repeat, _) = c.emit_loop_s(line);
    get(c, index, line);
    get(c, n, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_br_if(1, line);
    read(c, result, index, item, line);
    field(c, item, 2, line);
    set(c, value, line);
    write(c, result, index, value, line);
    increment(c, index, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(repeat);
    c.emit_end(line);
    c.patch_block(done);
    get(c, result, line);
}

/// Return the `n` smallest items from data. Stack: `[n, data]` -> `[array]`.
pub fn emit_nsmallest(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_n_of(chunks, current, argc, false, line);
}

/// Return the `n` largest items from data. Stack: `[n, data]` -> `[array]`.
pub fn emit_nlargest(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_n_of(chunks, current, argc, true, line);
}
