use vybe_compiler::primitives::{
    callable, class_context, class_slots, collections, heap, loops, references,
};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;
use vybe_runtime::opcode::heaptype::HT_EXTERN;

use class_slots::{ClassSlot, Dest, ObjSource, PlainNames, ResolvedSlot, ValueSource};

const HEAP_NUMERIC_COMPARATOR_CHUNK: &str = "__go_container_heap_numeric_compare";

pub fn emit_helper(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    match name {
        "go.container.heap.Init" if argc == 1 => emit_heap_init(chunks, current, line),
        "go.container.heap.Fix" if argc == 2 => emit_heap_fix(chunks, current, line),
        "go.container.heap.Pop" if argc == 1 => emit_heap_pop(chunks, current, line),
        "go.container.heap.Remove" if argc == 2 => emit_heap_remove(chunks, current, line),
        "go.container.heap.remove_prepare" if argc == 2 => {
            emit_heap_remove_prepare(chunks, current, line)
        }
        "go.container.list.New" if argc == 0 => emit_list_new(&mut chunks[current], line),
        "go.container.list.List.Init" if argc == 1 => emit_list_init(chunks, current, line),
        "go.container.list.List.Len" if argc == 1 => emit_get_i32(chunks, current, "len", line),
        "go.container.list.List.Front" if argc == 1 => {
            emit_get_field(chunks, current, "front", line)
        }
        "go.container.list.List.Back" if argc == 1 => emit_get_field(chunks, current, "back", line),
        "go.container.list.List.PushFront" if argc == 2 => {
            emit_list_push(chunks, current, true, line)
        }
        "go.container.list.List.PushBack" if argc == 2 => {
            emit_list_push(chunks, current, false, line)
        }
        "go.container.list.List.InsertBefore" if argc == 3 => {
            emit_list_insert(chunks, current, true, line)
        }
        "go.container.list.List.InsertAfter" if argc == 3 => {
            emit_list_insert(chunks, current, false, line)
        }
        "go.container.list.List.Remove" if argc == 2 => emit_list_remove(chunks, current, line),
        "go.container.list.List.MoveToFront" if argc == 2 => {
            emit_list_move_to_end(chunks, current, true, line)
        }
        "go.container.list.List.MoveToBack" if argc == 2 => {
            emit_list_move_to_end(chunks, current, false, line)
        }
        "go.container.list.List.MoveBefore" if argc == 3 => {
            emit_list_move(chunks, current, true, line)
        }
        "go.container.list.List.MoveAfter" if argc == 3 => {
            emit_list_move(chunks, current, false, line)
        }
        "go.container.list.List.PushBackList" if argc == 2 => {
            emit_list_copy(chunks, current, false, line)
        }
        "go.container.list.List.PushFrontList" if argc == 2 => {
            emit_list_copy(chunks, current, true, line)
        }
        "go.container.list.Element.Next" if argc == 1 => {
            emit_element_next_prev(chunks, current, true, line)
        }
        "go.container.list.Element.Prev" if argc == 1 => {
            emit_element_next_prev(chunks, current, false, line)
        }
        "go.container.ring.New" if argc == 1 => emit_ring_new(chunks, current, line),
        "go.container.ring.Ring.Next" if argc == 1 => emit_get_field(chunks, current, "next", line),
        "go.container.ring.Ring.Prev" if argc == 1 => emit_get_field(chunks, current, "prev", line),
        "go.container.ring.Ring.Len" if argc == 1 => emit_ring_len(chunks, current, line),
        "go.container.ring.Ring.Move" if argc == 2 => emit_ring_move(chunks, current, line),
        "go.container.ring.Ring.Do" if argc == 2 => emit_ring_do(chunks, current, line),
        "go.container.ring.Ring.Link" if argc == 2 => emit_ring_link(chunks, current, line),
        "go.container.ring.Ring.Unlink" if argc == 2 => emit_ring_unlink(chunks, current, line),
        "go.container.ring.Ring.Value.Get" if argc == 1 => {
            emit_public_value_get(chunks, current, line)
        }
        "go.container.ring.Ring.Value.Set" if argc == 2 => {
            emit_public_value_set(chunks, current, line)
        }
        _ => return false,
    }
    true
}

fn emit_ref_array_load(chunks: &mut [Chunk], current: usize, ref_slot: u16, dest: u16, line: u32) {
    lget(chunks, current, ref_slot, line);
    references::emit_cell_load(chunks, current, line);
    lset(chunks, current, dest, line);
}

fn ensure_heap_numeric_comparator(chunks: &mut Vec<Chunk>) -> usize {
    if let Some(idx) = chunks
        .iter()
        .position(|chunk| chunk.name == HEAP_NUMERIC_COMPARATOR_CHUNK)
    {
        return idx;
    }

    let mut comparator = Chunk::new(HEAP_NUMERIC_COMPARATOR_CHUNK);
    comparator.arity = 2;
    comparator.local_count = 2;
    comparator.emit_op_u16(Op::LOCAL_GET, 0, 0);
    comparator.emit_op_u16(Op::LOCAL_GET, 1, 0);
    comparator.emit_op(Op::F64_SUB, 0);
    comparator.emit_op(Op::RETURN, 0);
    chunks.push(comparator);
    chunks.len() - 1
}

fn emit_heap_sort(chunks: &mut Vec<Chunk>, current: usize, arr: u16, line: u32) {
    let comparator_idx = ensure_heap_numeric_comparator(chunks);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::REF_FUNC, comparator_idx as u16, line);
    chunks[current].emit(0, line);
    collections::emit_sort_with_comparator(chunks, current, line);
}

fn emit_heap_init(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let href = base;
    let arr = base + 1;
    lset(chunks, current, href, line);
    emit_ref_array_load(chunks, current, href, arr, line);
    emit_heap_sort(chunks, current, arr, line);
}

fn emit_heap_fix(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    emit_heap_init(chunks, current, line);
}

fn emit_heap_pop(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let href = base;
    let arr = base + 1;
    lset(chunks, current, href, line);
    emit_ref_array_load(chunks, current, href, arr, line);
    lget(chunks, current, arr, line);
    heap::emit_pop(chunks, current, 1, line);
}

fn emit_heap_remove(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let href = base;
    let index = base + 1;
    let arr = base + 2;
    let removed = base + 3;
    lset(chunks, current, index, line);
    lset(chunks, current, href, line);
    emit_ref_array_load(chunks, current, href, arr, line);
    emit_heap_sort(chunks, current, arr, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(chunks, current, arr, line);
    lget(chunks, current, index, line);
    collections::emit_get(chunks, current, line);
    lset(chunks, current, removed, line);
    lget(chunks, current, arr, line);
    lget(chunks, current, index, line);
    chunks[current].emit_i32_const(1, line);
    let splice = chunks[current].add_import("ecma:array", "splice");
    chunks[current].emit_call(splice, 3, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_heap_sort(chunks, current, arr, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(chunks, current, removed, line);
}

fn emit_heap_remove_prepare(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(6);
    let href = base;
    let index = base + 1;
    let arr = base + 2;
    let len = base + 3;
    let last = base + 4;
    let tmp = base + 5;
    lset(chunks, current, index, line);
    lset(chunks, current, href, line);
    emit_ref_array_load(chunks, current, href, arr, line);
    emit_heap_sort(chunks, current, arr, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(chunks, current, arr, line);
    collections::emit_len(chunks, current, line);
    lset(chunks, current, len, line);
    lget(chunks, current, len, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    lset(chunks, current, last, line);

    lget(chunks, current, index, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    lget(chunks, current, index, line);
    lget(chunks, current, len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_AND, line);
    lget(chunks, current, index, line);
    lget(chunks, current, last, line);
    chunks[current].emit_op(Op::I32_NE, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);
    emit_array_swap(chunks, current, arr, index, last, tmp, line);
    chunks[current].emit_end(line);
    null(&mut chunks[current], line);
}

fn emit_array_swap(
    chunks: &mut [Chunk],
    current: usize,
    arr: u16,
    i: u16,
    j: u16,
    tmp: u16,
    line: u32,
) {
    lget(chunks, current, arr, line);
    lget(chunks, current, i, line);
    collections::emit_get(chunks, current, line);
    lset(chunks, current, tmp, line);
    lget(chunks, current, arr, line);
    lget(chunks, current, i, line);
    lget(chunks, current, arr, line);
    lget(chunks, current, j, line);
    collections::emit_get(chunks, current, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(chunks, current, arr, line);
    lget(chunks, current, j, line);
    lget(chunks, current, tmp, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_list_new(chunk: &mut Chunk, line: u32) {
    class_slots::emit_class_construct(
        chunk,
        "__goList",
        &[
            (slot("front"), ValueSource::Null),
            (slot("back"), ValueSource::Null),
            (slot("len"), ValueSource::ConstI32(0)),
        ],
        line,
    );
}

fn emit_list_element_from_stack(chunk: &mut Chunk, line: u32) {
    class_slots::emit_class_construct(
        chunk,
        "__goListElement",
        &[
            (public_slot("Value"), ValueSource::Stack),
            (slot("next"), ValueSource::Null),
            (slot("prev"), ValueSource::Null),
            (slot("list"), ValueSource::Null),
        ],
        line,
    );
}

fn emit_list_init(chunks: &mut [Chunk], current: usize, line: u32) {
    let list = chunks[current].alloc_scratch(1);
    lset(chunks, current, list, line);
    set_null(&mut chunks[current], list, "front", line);
    set_null(&mut chunks[current], list, "back", line);
    set_i32(&mut chunks[current], list, "len", 0, line);
    lget(chunks, current, list, line);
}

fn emit_get_i32(chunks: &mut [Chunk], current: usize, field: &str, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    lset(chunks, current, value, line);
    get_to(&mut chunks[current], value, field, value, line);
    lget(chunks, current, value, line);
}

fn emit_get_field(chunks: &mut [Chunk], current: usize, field: &str, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    lset(chunks, current, value, line);
    get_stack(&mut chunks[current], value, field, line);
}

fn emit_element_next_prev(chunks: &mut [Chunk], current: usize, next: bool, line: u32) {
    let elem = chunks[current].alloc_scratch(1);
    lset(chunks, current, elem, line);
    lget(chunks, current, elem, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    null(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get_stack(
        &mut chunks[current],
        elem,
        if next { "next" } else { "prev" },
        line,
    );
    chunks[current].emit_end(line);
}

fn emit_list_push(chunks: &mut [Chunk], current: usize, front: bool, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let list = base;
    let value = base + 1;
    let elem = base + 2;
    let mark = base + 3;
    lset(chunks, current, value, line);
    lset(chunks, current, list, line);
    lget(chunks, current, value, line);
    emit_list_element_from_stack(&mut chunks[current], line);
    lset(chunks, current, elem, line);
    if front {
        get_to(&mut chunks[current], list, "front", mark, line);
        emit_insert_between_locals(chunks, current, list, elem, None, Some(mark), line);
    } else {
        get_to(&mut chunks[current], list, "back", mark, line);
        emit_insert_between_locals(chunks, current, list, elem, Some(mark), None, line);
    }
}

fn emit_list_insert(chunks: &mut [Chunk], current: usize, before: bool, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let list = base;
    let value = base + 1;
    let mark = base + 2;
    let elem = base + 3;
    let neighbor = base + 4;
    lset(chunks, current, mark, line);
    lset(chunks, current, value, line);
    lset(chunks, current, list, line);

    emit_mark_in_list(chunks, current, mark, list, line);
    chunks[current].emit_if(line);
    lget(chunks, current, value, line);
    emit_list_element_from_stack(&mut chunks[current], line);
    lset(chunks, current, elem, line);
    if before {
        get_to(&mut chunks[current], mark, "prev", neighbor, line);
        emit_insert_between_locals(
            chunks,
            current,
            list,
            elem,
            Some(neighbor),
            Some(mark),
            line,
        );
    } else {
        get_to(&mut chunks[current], mark, "next", neighbor, line);
        emit_insert_between_locals(
            chunks,
            current,
            list,
            elem,
            Some(mark),
            Some(neighbor),
            line,
        );
    }
    chunks[current].emit_else(line);
    null(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

fn emit_list_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let list = base;
    let elem = base + 1;
    let value = base + 2;
    let prev = base + 3;
    let next = base + 4;
    lset(chunks, current, elem, line);
    lset(chunks, current, list, line);

    emit_mark_in_list(chunks, current, elem, list, line);
    chunks[current].emit_if(line);
    get_to(&mut chunks[current], elem, "Value", value, line);
    get_to(&mut chunks[current], elem, "prev", prev, line);
    get_to(&mut chunks[current], elem, "next", next, line);
    emit_detach_known_locals(chunks, current, list, elem, prev, next, line);
    set_null(&mut chunks[current], elem, "list", line);
    set_null(&mut chunks[current], elem, "next", line);
    set_null(&mut chunks[current], elem, "prev", line);
    lget(chunks, current, value, line);
    chunks[current].emit_else(line);
    null(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

fn emit_list_move_to_end(chunks: &mut [Chunk], current: usize, front: bool, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let list = base;
    let elem = base + 1;
    let prev = base + 2;
    let next = base + 3;
    let mark = base + 4;
    lset(chunks, current, elem, line);
    lset(chunks, current, list, line);

    emit_mark_in_list(chunks, current, elem, list, line);
    chunks[current].emit_if(line);
    get_to(&mut chunks[current], elem, "prev", prev, line);
    get_to(&mut chunks[current], elem, "next", next, line);
    emit_detach_known_locals(chunks, current, list, elem, prev, next, line);
    if front {
        get_to(&mut chunks[current], list, "front", mark, line);
        emit_insert_between_locals(chunks, current, list, elem, None, Some(mark), line);
    } else {
        get_to(&mut chunks[current], list, "back", mark, line);
        emit_insert_between_locals(chunks, current, list, elem, Some(mark), None, line);
    }
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    null(&mut chunks[current], line);
}

fn emit_list_move(chunks: &mut [Chunk], current: usize, before: bool, line: u32) {
    let base = chunks[current].alloc_scratch(6);
    let list = base;
    let elem = base + 1;
    let mark = base + 2;
    let prev = base + 3;
    let next = base + 4;
    let neighbor = base + 5;
    lset(chunks, current, mark, line);
    lset(chunks, current, elem, line);
    lset(chunks, current, list, line);

    emit_mark_in_list(chunks, current, elem, list, line);
    emit_mark_in_list(chunks, current, mark, list, line);
    chunks[current].emit_op(Op::I32_AND, line);
    lget(chunks, current, elem, line);
    lget(chunks, current, mark, line);
    chunks[current].emit_op(Op::REF_EQ, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);
    get_to(&mut chunks[current], elem, "prev", prev, line);
    get_to(&mut chunks[current], elem, "next", next, line);
    if before {
        lget(chunks, current, next, line);
        lget(chunks, current, mark, line);
    } else {
        lget(chunks, current, prev, line);
        lget(chunks, current, mark, line);
    }
    chunks[current].emit_op(Op::REF_EQ, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    emit_detach_known_locals(chunks, current, list, elem, prev, next, line);
    if before {
        get_to(&mut chunks[current], mark, "prev", neighbor, line);
        emit_insert_between_locals(
            chunks,
            current,
            list,
            elem,
            Some(neighbor),
            Some(mark),
            line,
        );
    } else {
        get_to(&mut chunks[current], mark, "next", neighbor, line);
        emit_insert_between_locals(
            chunks,
            current,
            list,
            elem,
            Some(mark),
            Some(neighbor),
            line,
        );
    }
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    null(&mut chunks[current], line);
}

fn emit_list_copy(chunks: &mut [Chunk], current: usize, front: bool, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let list = base;
    let other = base + 1;
    let cur = base + 2;
    let value = base + 3;
    let tmp = base + 4;
    lset(chunks, current, other, line);
    lset(chunks, current, list, line);
    get_to(
        &mut chunks[current],
        other,
        if front { "back" } else { "front" },
        cur,
        line,
    );

    let state = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, cur, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);

    get_to(&mut chunks[current], cur, "Value", value, line);
    lget(chunks, current, list, line);
    lget(chunks, current, value, line);
    if front {
        emit_list_push(chunks, current, true, line);
    } else {
        emit_list_push(chunks, current, false, line);
    }
    chunks[current].emit_op(Op::DROP, line);
    get_to(
        &mut chunks[current],
        cur,
        if front { "prev" } else { "next" },
        tmp,
        line,
    );
    lget(chunks, current, tmp, line);
    lset(chunks, current, cur, line);
    loops::emit_loop_end(chunks, current, state, line);
    null(&mut chunks[current], line);
}

fn emit_insert_between_locals(
    chunks: &mut [Chunk],
    current: usize,
    list: u16,
    elem: u16,
    prev: Option<u16>,
    next: Option<u16>,
    line: u32,
) {
    set_local(&mut chunks[current], elem, "list", list, line);
    match prev {
        Some(prev) => set_local(&mut chunks[current], elem, "prev", prev, line),
        None => set_null(&mut chunks[current], elem, "prev", line),
    }
    match next {
        Some(next) => set_local(&mut chunks[current], elem, "next", next, line),
        None => set_null(&mut chunks[current], elem, "next", line),
    }

    emit_set_prev_next_or_front(chunks, current, list, elem, prev, line);
    emit_set_next_prev_or_back(chunks, current, list, elem, next, line);

    get_stack(&mut chunks[current], list, "len", line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set_stack(&mut chunks[current], list, "len", line);
    lget(chunks, current, elem, line);
}

fn emit_set_prev_next_or_front(
    chunks: &mut [Chunk],
    current: usize,
    list: u16,
    elem: u16,
    prev: Option<u16>,
    line: u32,
) {
    if let Some(prev) = prev {
        lget(chunks, current, prev, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);
        set_local(&mut chunks[current], list, "front", elem, line);
        chunks[current].emit_else(line);
        set_local(&mut chunks[current], prev, "next", elem, line);
        chunks[current].emit_end(line);
    } else {
        set_local(&mut chunks[current], list, "front", elem, line);
    }
}

fn emit_set_next_prev_or_back(
    chunks: &mut [Chunk],
    current: usize,
    list: u16,
    elem: u16,
    next: Option<u16>,
    line: u32,
) {
    if let Some(next) = next {
        lget(chunks, current, next, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);
        set_local(&mut chunks[current], list, "back", elem, line);
        chunks[current].emit_else(line);
        set_local(&mut chunks[current], next, "prev", elem, line);
        chunks[current].emit_end(line);
    } else {
        set_local(&mut chunks[current], list, "back", elem, line);
    }
}

fn emit_detach_known_locals(
    chunks: &mut [Chunk],
    current: usize,
    list: u16,
    _elem: u16,
    prev: u16,
    next: u16,
    line: u32,
) {
    emit_set_prev_next_or_front(chunks, current, list, next, Some(prev), line);
    emit_set_next_prev_or_back(chunks, current, list, prev, Some(next), line);
    get_stack(&mut chunks[current], list, "len", line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    set_stack(&mut chunks[current], list, "len", line);
}

fn emit_mark_in_list(chunks: &mut [Chunk], current: usize, mark: u16, list: u16, line: u32) {
    lget(chunks, current, mark, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    get_stack(&mut chunks[current], mark, "list", line);
    lget(chunks, current, list, line);
    chunks[current].emit_op(Op::REF_EQ, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

fn emit_ring_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let n = base;
    let first = base + 1;
    let prev = base + 2;
    let node = base + 3;
    lset(chunks, current, n, line);
    lget(chunks, current, n, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LE_S, line);
    chunks[current].emit_if(line);
    null(&mut chunks[current], line);
    chunks[current].emit_else(line);
    emit_ring_node(&mut chunks[current], line);
    lset(chunks, current, first, line);
    lget(chunks, current, first, line);
    lset(chunks, current, prev, line);
    lget(chunks, current, n, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    lset(chunks, current, n, line);
    let state = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, n, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);
    emit_ring_node(&mut chunks[current], line);
    lset(chunks, current, node, line);
    set_local(&mut chunks[current], prev, "next", node, line);
    set_local(&mut chunks[current], node, "prev", prev, line);
    lget(chunks, current, node, line);
    lset(chunks, current, prev, line);
    lget(chunks, current, n, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    lset(chunks, current, n, line);
    loops::emit_loop_end(chunks, current, state, line);
    set_local(&mut chunks[current], prev, "next", first, line);
    set_local(&mut chunks[current], first, "prev", prev, line);
    lget(chunks, current, first, line);
    chunks[current].emit_end(line);
}

fn emit_ring_node(chunk: &mut Chunk, line: u32) {
    class_slots::emit_class_construct(
        chunk,
        "__goRing",
        &[
            (public_slot("Value"), ValueSource::Null),
            (slot("next"), ValueSource::Null),
            (slot("prev"), ValueSource::Null),
        ],
        line,
    );
}

fn emit_ring_len(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let ring = base;
    let cur = base + 1;
    let n = base + 2;
    lset(chunks, current, ring, line);
    lget(chunks, current, ring, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(1, line);
    lset(chunks, current, n, line);
    get_to(&mut chunks[current], ring, "next", cur, line);
    let state = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, cur, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    lget(chunks, current, cur, line);
    lget(chunks, current, ring, line);
    chunks[current].emit_op(Op::REF_EQ, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op(Op::I32_AND, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);
    lget(chunks, current, n, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, n, line);
    get_to(&mut chunks[current], cur, "next", cur, line);
    loops::emit_loop_end(chunks, current, state, line);
    lget(chunks, current, n, line);
    chunks[current].emit_end(line);
}

fn emit_ring_move(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let ring = base;
    let n = base + 1;
    let cur = base + 2;
    let original = base + 3;
    lset(chunks, current, n, line);
    lset(chunks, current, ring, line);
    lget(chunks, current, n, line);
    lset(chunks, current, original, line);
    set_local(&mut chunks[current], ring, "Value", original, line);
    lget(chunks, current, ring, line);
    lset(chunks, current, cur, line);
    let state = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, n, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_NE, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);
    lget(chunks, current, n, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    get_to(&mut chunks[current], cur, "next", cur, line);
    lget(chunks, current, n, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    lset(chunks, current, n, line);
    chunks[current].emit_else(line);
    get_to(&mut chunks[current], cur, "prev", cur, line);
    lget(chunks, current, n, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, n, line);
    chunks[current].emit_end(line);
    loops::emit_loop_end(chunks, current, state, line);
    lget(chunks, current, cur, line);
}

fn emit_ring_do(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let ring = base;
    let func = base + 1;
    let cur = base + 2;
    lset(chunks, current, func, line);
    lset(chunks, current, ring, line);
    lget(chunks, current, ring, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    null(&mut chunks[current], line);
    chunks[current].emit_else(line);
    lget(chunks, current, ring, line);
    lset(chunks, current, cur, line);
    let state = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, cur, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);
    lget(chunks, current, func, line);
    let abi = class_context::module_receiver_abi(chunks);
    let recv = callable::emit_callback_receiver(&mut chunks[current], abi, line);
    get_stack(&mut chunks[current], cur, "Value", line);
    callable::emit_direct_invoke_chunk(&mut chunks[current], 1 + recv, line);
    chunks[current].emit_op(Op::DROP, line);
    get_to(&mut chunks[current], cur, "next", cur, line);
    lget(chunks, current, cur, line);
    lget(chunks, current, ring, line);
    chunks[current].emit_op(Op::REF_EQ, line);
    chunks[current].emit_if(line);
    null(&mut chunks[current], line);
    lset(chunks, current, cur, line);
    chunks[current].emit_end(line);
    loops::emit_loop_end(chunks, current, state, line);
    null(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

fn emit_public_value_get(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj = chunks[current].alloc_scratch(1);
    lset(chunks, current, obj, line);
    get_stack(&mut chunks[current], obj, "Value", line);
}

fn emit_public_value_set(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let obj = base;
    let value = base + 1;
    lset(chunks, current, value, line);
    lset(chunks, current, obj, line);
    lget(chunks, current, value, line);
    set_stack(&mut chunks[current], obj, "Value", line);
    lget(chunks, current, value, line);
}

fn emit_ring_link(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let ring = base;
    let other = base + 1;
    let rn = base + 2;
    let op = base + 3;
    lset(chunks, current, other, line);
    lset(chunks, current, ring, line);
    lget(chunks, current, ring, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    lget(chunks, current, other, line);
    chunks[current].emit_else(line);
    lget(chunks, current, other, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    get_stack(&mut chunks[current], ring, "next", line);
    chunks[current].emit_else(line);
    get_to(&mut chunks[current], ring, "next", rn, line);
    get_to(&mut chunks[current], other, "prev", op, line);
    set_local(&mut chunks[current], ring, "next", other, line);
    set_local(&mut chunks[current], other, "prev", ring, line);
    set_local(&mut chunks[current], op, "next", rn, line);
    set_local(&mut chunks[current], rn, "prev", op, line);
    lget(chunks, current, rn, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_ring_unlink(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let ring = base;
    let n = base + 1;
    let first = base + 2;
    let last = base + 3;
    lset(chunks, current, n, line);
    lset(chunks, current, ring, line);
    lget(chunks, current, ring, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    lget(chunks, current, n, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LE_S, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    null(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get_to(&mut chunks[current], ring, "next", first, line);
    lget(chunks, current, first, line);
    lset(chunks, current, last, line);
    lget(chunks, current, n, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    lset(chunks, current, n, line);
    let state = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, n, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    get_stack(&mut chunks[current], last, "next", line);
    lget(chunks, current, ring, line);
    chunks[current].emit_op(Op::REF_EQ, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op(Op::I32_AND, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);
    get_to(&mut chunks[current], last, "next", last, line);
    lget(chunks, current, n, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    lset(chunks, current, n, line);
    loops::emit_loop_end(chunks, current, state, line);
    get_stack(&mut chunks[current], last, "next", line);
    set_stack(&mut chunks[current], ring, "next", line);
    set_local(&mut chunks[current], ring, "prev", ring, line);
    set_local(&mut chunks[current], first, "prev", last, line);
    set_local(&mut chunks[current], last, "next", first, line);
    lget(chunks, current, first, line);
    chunks[current].emit_end(line);
}

fn slot(field: &str) -> ResolvedSlot {
    class_slots::resolve(&ClassSlot::internal(field), &PlainNames)
}

fn public_slot(field: &str) -> ResolvedSlot {
    class_slots::resolve(&ClassSlot::instance(field), &PlainNames)
}

fn get_stack(chunk: &mut Chunk, obj: u16, field: &str, line: u32) {
    let slot = if field == "Value" {
        public_slot(field)
    } else {
        slot(field)
    };
    class_slots::emit_class_get(chunk, ObjSource::Local(obj), &slot, Dest::Stack, line);
}

fn get_to(chunk: &mut Chunk, obj: u16, field: &str, dest: u16, line: u32) {
    let slot = if field == "Value" {
        public_slot(field)
    } else {
        slot(field)
    };
    class_slots::emit_class_get(chunk, ObjSource::Local(obj), &slot, Dest::Local(dest), line);
}

fn set_stack(chunk: &mut Chunk, obj: u16, field: &str, line: u32) {
    let slot = if field == "Value" {
        public_slot(field)
    } else {
        slot(field)
    };
    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(obj),
        &slot,
        ValueSource::Stack,
        line,
    );
}

fn set_local(chunk: &mut Chunk, obj: u16, field: &str, value: u16, line: u32) {
    let slot = if field == "Value" {
        public_slot(field)
    } else {
        slot(field)
    };
    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(obj),
        &slot,
        ValueSource::Local(value),
        line,
    );
}

fn set_null(chunk: &mut Chunk, obj: u16, field: &str, line: u32) {
    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(obj),
        &slot(field),
        ValueSource::Null,
        line,
    );
}

fn set_i32(chunk: &mut Chunk, obj: u16, field: &str, value: i32, line: u32) {
    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(obj),
        &slot(field),
        ValueSource::ConstI32(value),
        line,
    );
}

fn null(chunk: &mut Chunk, line: u32) {
    chunk.emit_ref_null(HT_EXTERN, line);
}

fn lget(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}
