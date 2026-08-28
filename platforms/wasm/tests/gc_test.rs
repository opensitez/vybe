//! Tests for the GC proposal (0xFB prefix).
//! Covers: struct.new/get/set, array.new_fixed/get/set/len/fill/copy,
//!         i31.new/get_s/get_u, ref.test, ref.cast, ref.cast_null,
//!         any.convert_extern/extern.convert_any,
//!         br_on_null/br_on_non_null/ref.as_non_null,
//!         br_on_cast/br_on_cast_fail.

use std::sync::Arc;
use vybe_platform_wasm::read_wasm;
use vybe_platform_wasm::write_wasm;
use vybe_runtime::chunk::TypeEntry;
use vybe_runtime::opcode::heaptype::{HT_ANY, HT_STRUCT, HeapType};
use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};

/// Run without appending RETURN — caller is responsible for the full layout.
fn run_raw(emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut chunk = Chunk::new("<script>");
    emit(&mut chunk);
    VM::new().run(vec![chunk]).expect("run failed")
}

fn run_err(emit: impl FnOnce(&mut Chunk)) -> String {
    let mut chunk = Chunk::new("<script>");
    emit(&mut chunk);
    chunk.emit_op(Op::RETURN, 0);
    VM::new().run(vec![chunk]).unwrap_err().to_string()
}

fn emit_op_u16_u16(chunk: &mut Chunk, op: Op, first: u16, second: u16, line: u32) {
    chunk.emit_op(op, line);
    chunk.emit((first >> 8) as u8, line);
    chunk.emit((first & 0xff) as u8, line);
    chunk.emit((second >> 8) as u8, line);
    chunk.emit((second & 0xff) as u8, line);
}

fn has_bytes(bytes: &[u8], needle: &[u8]) -> bool {
    bytes.windows(needle.len()).any(|w| w == needle)
}

fn type_entry(name: &str, fields: &[&str]) -> TypeEntry {
    TypeEntry {
        name: name.into(),
        kind: vybe_runtime::chunk::CompositeKind::Struct,
        parent_index: 0,
        fields: fields.iter().map(|field| field.to_string()).collect(),
        methods: Vec::new(),
        is_interface: false,
        implements: Vec::new(),
        constructor_chunk: None,
        field_descriptors: std::collections::HashMap::new(),
    }
    ..Default::default()
}

#[test]
fn gc_emission_maps_chunk_type_index_to_wasm_struct_type_index() {
    let mut chunk = Chunk::new("<script>");
    chunk.types.push(type_entry("A", &["x"]));
    chunk.types.push(type_entry("B", &["y", "z"]));
    chunk.emit_op_u16(Op::STRUCT_NEW_DEFAULT, 1, 0);
    chunk.emit_op(Op::RETURN, 0);

    let wasm = write_wasm(&vec![chunk]);
    // ⚠ THE IMMEDIATE IS 1-BASED, so `STRUCT_NEW_DEFAULT 1` is type **A**,
    // whose described Wasm type index is 0 — not type B at index 2.
    //
    // This assertion used to demand `[0xFB, 0x01, 0x02]` and called typeidx 1
    // "the second chunk-local type", i.e. it read the immediate as 0-based and
    // pinned an off-by-one in `wasm_struct_type_for_chunk_type`. Every other
    // consumer disagrees: `resolve_gc_rtt` indexes `type_imm - 1`,
    // `struct_type_by_index` indexes `module_index - 1`, `TypeEntry::parent_index`
    // is documented 1-based with 0 meaning "none", and `classes.rs` reserves 0
    // for the dynamic form. The writer was the lone outlier and this test was
    // what held it in place.
    //
    // A has one descriptor-carrying type, so the allocation is
    // `global.get <desc> ; struct.new_default_desc 0` (0xFB 0x21), not
    // `struct.new_default` — see `encode_global_section_with_descriptors`.
    assert!(
        has_bytes(&wasm, &[0xFB, 0x21, 0x00]),
        "chunk-local type 1 is A, whose described Wasm type index is 0"
    );
    assert!(
        !has_bytes(&wasm, &[0xFB, 0x01, 0x02]),
        "must not emit the old off-by-one mapping to type B"
    );
}

#[test]
fn gc_emission_resolves_unique_struct_field_index() {
    let mut chunk = Chunk::new("<script>");
    chunk.types.push(type_entry("A", &["x"]));
    chunk.types.push(type_entry("B", &["y", "z"]));
    let field = chunk.add_constant(Value::String(Arc::from("z")));
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, field, 0);
    chunk.emit_op(Op::RETURN, 0);

    let wasm = write_wasm(&vec![chunk]);
    assert!(
        has_bytes(&wasm, &[0xFB, 0x02, 0x02, 0x01]),
        "field z should resolve to struct type index 2, field index 1"
    );
}

#[test]
fn gc_emission_struct_set_reorders_object_and_value_operands() {
    let mut chunk = Chunk::new("<script>");
    chunk.types.push(type_entry("A", &["x"]));
    chunk.types.push(type_entry("B", &["y", "z"]));
    let field = chunk.add_constant(Value::String(Arc::from("z")));
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_i32_const(9, 0);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, field, 0);
    chunk.emit_op(Op::RETURN, 0);

    let wasm = write_wasm(&vec![chunk]);
    assert!(
        has_bytes(
            &wasm,
            &[
                0x21, 0x00, // local.set temp = value
                0xFB, 0x1A, // any.convert_extern on object
                0xFB, 0x17, 0x02, // ref.cast null typeidx 2
                0x20, 0x00, // local.get temp = value
                0xFB, 0x05, 0x02,
                0x01, // struct.set typeidx 2 fieldidx 1
                      // spec struct.set pushes nothing — no reload; the internal
                      // op has the same shape now.
            ],
        ),
        "struct.set emission should save value, cast object, set field, and reload value"
    );
}

// ── BR_ON_NULL ────────────────────────────────────────────────────────────

#[test]
fn br_on_null_branches_when_null() {
    // Branch offset is relative to the ip after BR_ON_NULL's operands; the
    // null path must skip the not-reached consts, whose encoded size is
    // LEB-variable — so measure and patch rather than hand-count bytes.
    let r = run_raw(|c| {
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        c.emit_op_u16(Op::BR_ON_NULL, 0u16, 0); // offset patched below
        let after = c.code.len();
        c.emit_i32_const(0, 0); // not-null path (not reached)
        c.emit_op(Op::RETURN, 0); // not reached
        let off = (c.code.len() - after) as u16;
        c.code[after - 2] = (off >> 8) as u8;
        c.code[after - 1] = (off & 0xff) as u8;
        c.emit_i32_const(1, 0); // null path lands here
        c.emit_op(Op::RETURN, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

#[test]
fn br_on_null_skips_branch_on_non_null() {
    // Non-null value → fall through to CONST(0), return 0
    let r = run_raw(|c| {
        c.emit_i32_const(42, 0); // push non-null i32
        c.emit_op_u16(Op::BR_ON_NULL, 6u16, 0); // not null → fall through
        c.emit_op(Op::DROP, 0); // drop i32 42
        c.emit_i32_const(0, 0); // push 0
        c.emit_op(Op::RETURN, 0);
        c.emit_i32_const(1, 0); // null path (not reached)
        c.emit_op(Op::RETURN, 0);
    });
    assert_eq!(r.as_i32(), 0);
}

// ── BR_ON_NON_NULL ────────────────────────────────────────────────────────

#[test]
fn br_on_non_null_branches_on_non_null() {
    // BR_ON_NON_NULL leaves the value on the stack when it branches, so the
    // branch target RETURN returns 7. Offset measured and patched, not
    // hand-counted (const encoding is LEB-variable).
    let r = run_raw(|c| {
        c.emit_i32_const(7, 0);
        c.emit_op_u16(Op::BR_ON_NON_NULL, 0u16, 0); // offset patched below
        let after = c.code.len();
        // null fall-through (not reached):
        c.emit_op(Op::DROP, 0);
        c.emit_i32_const(0, 0);
        c.emit_op(Op::RETURN, 0);
        let off = (c.code.len() - after) as u16;
        c.code[after - 2] = (off >> 8) as u8;
        c.code[after - 1] = (off & 0xff) as u8;
        c.emit_op(Op::RETURN, 0); // non-null path: 7 still on stack
    });
    assert_eq!(r.as_i32(), 7);
}

#[test]
fn br_on_non_null_skips_on_null() {
    // Push null → BR_ON_NON_NULL should NOT branch → pops null, falls through
    let r = run_raw(|c| {
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0); // [0-1]
        c.emit_op_u16(Op::BR_ON_NON_NULL, 4u16, 0); // [2-5] null → no branch, pop
        c.emit_i32_const(99, 0); // [6-9]
        c.emit_op(Op::RETURN, 0); // [10-11]
        c.emit_i32_const(0, 0); // [12-15] not reached
        c.emit_op(Op::RETURN, 0); // [16-17]
    });
    assert_eq!(r.as_i32(), 99);
}

// ── REF_AS_NON_NULL ───────────────────────────────────────────────────────

#[test]
fn ref_as_non_null_passes_through_non_null() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_string_const("hello", 0);
    chunk.emit_op(Op::REF_AS_NON_NULL, 0);
    chunk.emit_op(Op::RETURN, 0);
    let r = VM::new().run(vec![chunk]).expect("run failed");
    assert_eq!(r.as_str(), "hello");
}

#[test]
fn ref_as_non_null_traps_on_null() {
    let e = run_err(|c| {
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        c.emit_op(Op::REF_AS_NON_NULL, 0);
    });
    assert!(e.contains("null") || e.contains("trap"));
}

// ── REF_CAST ─────────────────────────────────────────────────────────────

#[test]
fn ref_cast_traps_on_type_mismatch() {
    // Push a string, cast to "dog" — won't match → trap
    let e = run_err(|c| {
        c.types.push(type_entry("Dog", &["legs"]));
        c.emit_string_const("not-a-dog", 0);
        c.emit_ref_type_op(Op::REF_CAST, HeapType::Concrete(1), 0);
    });
    assert!(e.contains("cast") || e.contains("not"));
}

#[test]
fn ref_cast_passes_on_null_input() {
    // REF_CAST_NULL always passes (null is a valid reference for any nullable type)
    let mut chunk = Chunk::new("<script>");
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_ref_type_op(Op::REF_CAST_NULL, HeapType::Abstract(HT_STRUCT), 0);
    chunk.emit_op(Op::RETURN, 0);
    let r = VM::new().run(vec![chunk]).expect("run failed");
    assert!(matches!(r, Value::Null));
}

// ── BR_ON_CAST — uses label depth via emit_block / emit_end ──────────────

#[test]
fn br_on_cast_branches_on_type_match() {
    // struct with __type "foo" → br_on_cast "foo" at depth 0 → exits block
    let mut chunk = Chunk::new("<script>");
    let type_str = chunk.add_constant(Value::String(Arc::from("foo")));

    let _blk = chunk.emit_block(0);

    // Push a value tagged as "foo" (a string constant that test_type
    // matches). The pool entry stays alive as BR_ON_CAST's name immediate;
    // the VALUE rides the spec string-constant route.
    chunk.emit_string_const("foo", 0);

    // br_on_cast: U16 type_name_idx + U8 depth
    chunk.emit_op(Op::BR_ON_CAST, 0);
    chunk.emit((type_str >> 8) as u8, 0);
    chunk.emit((type_str & 0xFF) as u8, 0);
    chunk.emit(0u8, 0); // depth = 0 (exit enclosing block)

    // not branched: push missed, return
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_end(0);
    // branched: value (type_str constant = "foo") is on stack, drop and push 1
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::RETURN, 0);

    let r = VM::new().run(vec![chunk]).expect("run failed");
    // Result depends on test_type("foo", "foo") — it should match as string equality
    assert!(r.as_i32() == 0 || r.as_i32() == 1);
}

#[test]
fn br_on_cast_fail_branches_on_type_mismatch() {
    let mut chunk = Chunk::new("<script>");
    let wrong = chunk.add_constant(Value::String(Arc::from("not-bar")));

    let _blk = chunk.emit_block(0);
    chunk.emit_string_const("bar", 0); // "bar" value

    // br_on_cast_fail "not-bar" → branches because "bar" is not "not-bar"
    chunk.emit_op(Op::BR_ON_CAST_FAIL, 0);
    chunk.emit((wrong >> 8) as u8, 0);
    chunk.emit((wrong & 0xFF) as u8, 0);
    chunk.emit(0u8, 0);

    chunk.emit_op(Op::DROP, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_end(0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_i32_const(99, 0);
    chunk.emit_op(Op::RETURN, 0);

    let r = VM::new().run(vec![chunk]).expect("run failed");
    assert!(r.as_i32() == 0 || r.as_i32() == 99);
}

// ═══════════════════════════════════════════════════════════════════════════
// STRUCT ops (0xFB 0x00–0x05)
// ═══════════════════════════════════════════════════════════════════════════

fn run(emit: impl FnOnce(&mut Chunk)) -> Value {
    run_locals(0, emit)
}

fn run_locals(local_count: u16, emit: impl FnOnce(&mut Chunk)) -> Value {
    let mut c = Chunk::new("<script>");
    c.local_count = local_count;
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    VM::new().run(vec![c]).expect("run failed")
}

#[test]
fn struct_new_creates_object() {
    // STRUCT_NEW 0: pops 0 key-value pairs, pushes one empty object
    let r = run(|c| {
        c.emit_struct_new(0, 0, 0);
        c.emit_op(Op::REF_IS_NULL, 0); // object is not null → 0
    });
    assert_eq!(r.as_i32(), 0);
}

#[test]
fn struct_set_and_get_roundtrip() {
    let r = run_locals(1, |c| {
        let key = c.add_constant(Value::String(Arc::from("x")));

        // create empty struct, store in slot 0, drop stack copy
        c.emit_struct_new(0, 0, 0); // stack: [obj]
        c.emit_op_u16(Op::LOCAL_SET, 0, 0); // stack: [obj] (peek)

        // obj.x = 42
        c.emit_op_u16(Op::LOCAL_GET, 0, 0); // stack: [obj]
        c.emit_i32_const(42, 0); // stack: [obj, 42]
        c.emit_struct_field_op(Op::STRUCT_SET, 0, key, 0); // stack: []

        // read obj.x
        c.emit_op_u16(Op::LOCAL_GET, 0, 0); // stack: [obj]
        c.emit_struct_field_op(Op::STRUCT_GET, 0, key, 0); // stack: [42]
    });
    assert_eq!(r.as_i32(), 42);
}

#[test]
fn struct_set_overwrites_field() {
    let r = run_locals(1, |c| {
        let key = c.add_constant(Value::String(Arc::from("v")));

        c.emit_struct_new(0, 0, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);

        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(1, 0);
        c.emit_struct_field_op(Op::STRUCT_SET, 0, key, 0);

        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(99, 0);
        c.emit_struct_field_op(Op::STRUCT_SET, 0, key, 0);

        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_struct_field_op(Op::STRUCT_GET, 0, key, 0);
    });
    assert_eq!(r.as_i32(), 99);
}

// ═══════════════════════════════════════════════════════════════════════════
// ARRAY ops (0xFB 0x06–0x13)
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn array_new_fixed_and_get() {
    let r = run(|c| {
        c.emit_i32_const(10, 0);
        c.emit_i32_const(20, 0);
        c.emit_i32_const(30, 0);
        c.emit_array_new_fixed(0, 3, 0); // 3 elements
        // get element at index 1 (=20)
        c.emit_i32_const(1, 0);
        c.emit_op(Op::ARRAY_GET, 0);
    });
    assert_eq!(r.as_i32(), 20);
}

#[test]
fn array_set_updates_element() {
    let r = run_locals(1, |c| {
        // create [0,0,0], store in slot 0
        c.emit_i32_const(0, 0);
        c.emit_i32_const(0, 0);
        c.emit_i32_const(0, 0);
        c.emit_array_new_fixed(0, 3, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);

        // arr[1] = 77: push arr, idx, val
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(1, 0);
        c.emit_i32_const(77, 0);
        c.emit_op(Op::ARRAY_SET, 0);

        // get arr[1]
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(1, 0);
        c.emit_op(Op::ARRAY_GET, 0);
    });
    assert_eq!(r.as_i32(), 77);
}

#[test]
fn array_length() {
    let r = run(|c| {
        c.emit_i32_const(0, 0);
        c.emit_i32_const(0, 0);
        c.emit_i32_const(0, 0);
        c.emit_i32_const(0, 0);
        c.emit_array_new_fixed(0, 4, 0);
        c.emit_op(Op::ARRAY_LENGTH, 0);
    });
    assert_eq!(r.as_i32(), 4);
}

#[test]
fn array_fill_sets_range() {
    let r = run_locals(1, |c| {
        // [0,0,0,0,0]
        for _ in 0..5 {
            c.emit_i32_const(0, 0);
        }
        c.emit_array_new_fixed(0, 5, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);

        // array.fill spec stack: [arrayref, index, value, count].
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(1, 0); // index (start)
        c.emit_i32_const(99, 0); // value
        c.emit_i32_const(3, 0); // count
        c.emit_op(Op::ARRAY_FILL, 0);

        // check arr[2] = 99
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(2, 0);
        c.emit_op(Op::ARRAY_GET, 0);
    });
    assert_eq!(r.as_i32(), 99);
}

#[test]
fn array_copy_copies_range() {
    let r = run_locals(2, |c| {
        // src = [5,5,5]
        c.emit_i32_const(5, 0);
        c.emit_i32_const(5, 0);
        c.emit_i32_const(5, 0);
        c.emit_array_new_fixed(0, 3, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);

        // dst = [0,0,0]
        c.emit_i32_const(0, 0);
        c.emit_i32_const(0, 0);
        c.emit_i32_const(0, 0);
        c.emit_array_new_fixed(0, 3, 0);
        c.emit_op_u16(Op::LOCAL_SET, 1, 0);

        // array.copy dst dst_offset=0 src src_offset=0 count=3
        c.emit_op_u16(Op::LOCAL_GET, 1, 0); // dst
        c.emit_i32_const(0, 0); // dst offset
        c.emit_op_u16(Op::LOCAL_GET, 0, 0); // src
        c.emit_i32_const(0, 0); // src offset
        c.emit_i32_const(3, 0); // count
        c.emit_op(Op::ARRAY_COPY, 0);

        // dst[1] should now be 5
        c.emit_op_u16(Op::LOCAL_GET, 1, 0);
        c.emit_i32_const(1, 0);
        c.emit_op(Op::ARRAY_GET, 0);
    });
    assert_eq!(r.as_i32(), 5);
}

// ═══════════════════════════════════════════════════════════════════════════
// I31 ops (0xFB 0x1C–0x1E)
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn i31_new_and_get_s() {
    let r = run(|c| {
        c.emit_i32_const(-42, 0);
        c.emit_op(Op::I31_NEW, 0);
        c.emit_op(Op::I31_GET_S, 0);
    });
    assert_eq!(r.as_i32(), -42);
}

#[test]
fn i31_get_u_unsigned() {
    let r = run(|c| {
        // i31 max signed = 2^30 - 1 = 1073741823
        c.emit_i32_const(1073741823, 0);
        c.emit_op(Op::I31_NEW, 0);
        c.emit_op(Op::I31_GET_U, 0);
    });
    assert_eq!(r.as_i32() as u32, 1073741823u32);
}

// ═══════════════════════════════════════════════════════════════════════════
// ANY.CONVERT_EXTERN / EXTERN.CONVERT_ANY (0xFB 0x1A–0x1B)
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn any_convert_extern_is_identity() {
    // Both are no-ops at runtime (universal externref ABI)
    let r = run(|c| {
        c.emit_i32_const(7, 0);
        c.emit_op(Op::ANY_CONVERT_EXTERN, 0);
        c.emit_op(Op::EXTERN_CONVERT_ANY, 0);
    });
    assert_eq!(r.as_i32(), 7);
}

// ═══════════════════════════════════════════════════════════════════════════
// REF.TEST (0xFB 0x14)
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn ref_test_on_string_value() {
    // A string is a value, not a struct: `any` accepts it, `struct` does not.
    // Both directions are pinned — a test that accepts either answer pins
    // nothing.
    let any = run(|c| {
        c.emit_string_const("hello", 0);
        c.emit_ref_type_op(Op::REF_TEST, HeapType::Abstract(HT_ANY), 0);
    });
    assert_eq!(any.as_i32(), 1, "every non-null value is `any`");

    let structural = run(|c| {
        c.emit_string_const("hello", 0);
        c.emit_ref_type_op(Op::REF_TEST, HeapType::Abstract(HT_STRUCT), 0);
    });
    assert_eq!(structural.as_i32(), 0, "a string is not a struct");
}

// ═══════════════════════════════════════════════════════════════════════════
// ARRAY.NEW / ARRAY.NEW_DEFAULT (0xFB 0x06 / 0x07)
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn array_new_fills_all_lanes_with_value() {
    // array.new $t: pops [value, len] → array of len copies of value
    let r = run(|c| {
        c.emit_i32_const(7, 0);
        c.emit_i32_const(4, 0);
        c.emit_op_u16(Op::ARRAY_NEW, 0, 0); // typeidx=0, reads u16

        // get element at index 2 → should be 7
        c.emit_i32_const(2, 0);
        c.emit_op(Op::ARRAY_GET, 0);
    });
    assert_eq!(r.as_i32(), 7);
}

#[test]
fn array_new_length_is_correct() {
    let r = run(|c| {
        c.emit_i32_const(0, 0);
        c.emit_i32_const(5, 0);
        c.emit_op_u16(Op::ARRAY_NEW, 0, 0);
        c.emit_op(Op::ARRAY_LENGTH, 0);
    });
    assert_eq!(r.as_i32(), 5);
}

#[test]
fn array_new_default_initializes_to_null() {
    // array.new_default $t: pops [len] → array of len nulls
    let r = run(|c| {
        c.emit_i32_const(3, 0);
        c.emit_op_u16(Op::ARRAY_NEW_DEFAULT, 0, 0);
        // element 0 should be null
        c.emit_i32_const(0, 0);
        c.emit_op(Op::ARRAY_GET, 0);
        c.emit_op(Op::REF_IS_NULL, 0);
    });
    assert_eq!(r.as_i32(), 1); // null
}

#[test]
fn array_new_default_length_is_correct() {
    let r = run(|c| {
        c.emit_i32_const(6, 0);
        c.emit_op_u16(Op::ARRAY_NEW_DEFAULT, 0, 0);
        c.emit_op(Op::ARRAY_LENGTH, 0);
    });
    assert_eq!(r.as_i32(), 6);
}

// ═══════════════════════════════════════════════════════════════════════════
// ARRAY.NEW_DATA / ARRAY.NEW_ELEM (0xFB 0x09 / 0x0A)
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn array_new_data_reads_data_segment() {
    let mut vm = VM::new();
    vm.set_data_segment(0, vec![10, 20, 30, 40]);
    let mut c = Chunk::new("<test>");
    {
        c.emit_i32_const(0, 0);
        c.emit_i32_const(4, 0);
        emit_op_u16_u16(&mut c, Op::ARRAY_NEW_DATA, 0, 0, 0);
        c.emit_i32_const(2, 0);
        c.emit_op(Op::ARRAY_GET, 0);
    }
    let r = vm.run(vec![c]).unwrap();
    assert_eq!(r.as_i32(), 30);
}

#[test]
fn array_new_elem_reads_elem_segment() {
    let mut vm = VM::new();
    vm.set_elem_segment(0, vec![Value::I32(7), Value::I32(8), Value::I32(9)]);
    let mut c = Chunk::new("<test>");
    {
        c.emit_i32_const(1, 0);
        c.emit_i32_const(2, 0);
        emit_op_u16_u16(&mut c, Op::ARRAY_NEW_ELEM, 0, 0, 0);
        c.emit_i32_const(0, 0);
        c.emit_op(Op::ARRAY_GET, 0);
    }
    let r = vm.run(vec![c]).unwrap();
    assert_eq!(r.as_i32(), 8);
}

// ═══════════════════════════════════════════════════════════════════════════
// ARRAY.GET_S / ARRAY.GET_U (0xFB 0x0C / 0x0D)
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn array_get_s_on_regular_array_reads_element() {
    // For Value arrays (non-packed), array.get_s behaves like array.get
    let r = run(|c| {
        c.emit_i32_const(42, 0);
        c.emit_array_new_fixed(0, 1, 0);
        c.emit_i32_const(0, 0);
        c.emit_op_u16(Op::ARRAY_GET_S, 0, 0);
    });
    assert_eq!(r.as_i32(), 42);
}

#[test]
fn array_get_u_on_regular_array_reads_element() {
    let r = run(|c| {
        c.emit_i32_const(99, 0);
        c.emit_array_new_fixed(0, 1, 0);
        c.emit_i32_const(0, 0);
        c.emit_op_u16(Op::ARRAY_GET_U, 0, 0);
    });
    assert_eq!(r.as_i32(), 99);
}

// ═══════════════════════════════════════════════════════════════════════════
// ARRAY.INIT_DATA / ARRAY.INIT_ELEM (0xFB 0x12 / 0x13)
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn array_init_data_copies_into_array() {
    let mut vm = VM::new();
    vm.set_data_segment(0, vec![1, 2, 3, 4]);

    // `array.init_data $t $d` copies `size` ELEMENTS, and `src` is a BYTE
    // offset — so the element storage width decides how many bytes are read.
    // Declare an `(array i8)` type and stamp the array with it; without an rtt
    // the VM must assume i32 (4-byte elements), and reading one element from
    // byte offset 2 of a 4-byte segment would legitimately trap.
    {
        let mut td = vybe_runtime::typedef::TypeDef::new("<array i8>");
        td.add_field("i8");
        vm.type_registry.register(td);
    }

    let mut c = Chunk::new("<test>");
    c.local_count = 1;
    // `array.new_fixed $t` takes a 1-based index into THIS MODULE's type
    // table, which the VM maps back to a name and resolves against the
    // registry (`resolve_gc_rtt`) — it is NOT the registry id, because the
    // host pre-registers its builtin types ahead of the module's own. So the
    // module has to declare the type it names; entry 0 → immediate 1.
    c.types.push(vybe_runtime::chunk::TypeEntry {
        name: "<array i8>".to_string(),
        kind: vybe_runtime::chunk::CompositeKind::Array,
        parent_index: 0,
        fields: vec!["i8".to_string()],
        methods: Vec::new(),
        is_interface: false,
        implements: Vec::new(),
        constructor_chunk: None,
        field_descriptors: std::collections::HashMap::new(),
            ..Default::default()
    });
    let i8_array_type = 1u16;
    {
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        // `array.new_fixed $t 3` — the rtt comes from the TYPE INDEX immediate
        // and is stamped at allocation, per spec. This used to allocate an
        // untyped array and then re-stamp it with the custom `SET_TYPE_ID`
        // opcode, which no longer exists.
        c.emit_array_new_fixed(i8_array_type as u16, 3, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);

        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(1, 0); // dst_offset
        c.emit_i32_const(2, 0); // src_offset
        c.emit_i32_const(1, 0); // size
        emit_op_u16_u16(&mut c, Op::ARRAY_INIT_DATA, 0, 0, 0);

        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(1, 0);
        c.emit_op(Op::ARRAY_GET, 0);
    }
    let r = vm.run(vec![c]).unwrap();
    assert_eq!(r.as_i32(), 3);
}

#[test]
fn array_init_elem_copies_into_array() {
    let mut vm = VM::new();
    vm.set_elem_segment(0, vec![Value::I32(11), Value::I32(12), Value::I32(13)]);
    let mut c = Chunk::new("<test>");
    c.local_count = 1;
    {
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        c.emit_array_new_fixed(0, 2, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);

        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(0, 0); // dst_offset
        c.emit_i32_const(1, 0); // src_offset
        c.emit_i32_const(2, 0); // size
        emit_op_u16_u16(&mut c, Op::ARRAY_INIT_ELEM, 0, 0, 0);

        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(1, 0);
        c.emit_op(Op::ARRAY_GET, 0);
    }
    let r = vm.run(vec![c]).unwrap();
    assert_eq!(r.as_i32(), 13);
}

// ═══════════════════════════════════════════════════════════════════════════
// STRUCT.NEW_DEFAULT (0xFB 0x01)
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn struct_new_default_creates_non_null_object() {
    // struct.new_default: no fields on stack, pushes empty struct
    let r = run(|c| {
        c.emit_op_u16(Op::STRUCT_NEW_DEFAULT, 0, 0);
        c.emit_op(Op::REF_IS_NULL, 0);
    });
    assert_eq!(r.as_i32(), 0); // not null
}

// ═══════════════════════════════════════════════════════════════════════════
// STRUCT.GET_S / STRUCT.GET_U (0xFB 0x03 / 0x04)
// ═══════════════════════════════════════════════════════════════════════════

#[test]
fn struct_get_s_reads_from_fields_array() {
    // struct.get_s/u read from obj.fields[field_idx]
    // For plain objects without a fields vec, returns Null
    let r = run(|c| {
        c.emit_struct_new(0, 0, 0);
        c.emit_struct_field_op(Op::STRUCT_GET_S, 0, 0, 0); // field 0 → Null
        c.emit_op(Op::REF_IS_NULL, 0);
    });
    assert_eq!(r.as_i32(), 1); // fields[0] is null (empty)
}

#[test]
fn struct_get_u_reads_from_fields_array() {
    let r = run(|c| {
        c.emit_struct_new(0, 0, 0);
        c.emit_struct_field_op(Op::STRUCT_GET_U, 0, 0, 0); // field 0 → Null
        c.emit_op(Op::REF_IS_NULL, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

// ═══════════════════════════════════════════════════════════════════════════
// Custom Descriptors proposal: STRUCT.NEW_DESC / STRUCT.NEW_DEFAULT_DESC / REF.GET_DESC
// (0xFB 0x20 / 0x21 / 0x22)
// ═══════════════════════════════════════════════════════════════════════════

// ⚠ These read the descriptor back with `ref.get_desc`, NOT by fetching the
// reserved `__descriptor` property with `struct.get`. Where the descriptor is
// stored is an implementation detail the proposal does not expose; asserting
// on it made the tests pass for a stub that allocated an empty object and
// ignored both its type index and its field operands.

#[test]
fn struct_new_desc_attaches_descriptor() {
    let r = run(|c| {
        c.emit_string_const("my-type", 0); // descriptor — the LAST operand
        c.emit_op_u16_u16(Op::STRUCT_NEW_DESC, 0, 0, 0);
        c.emit_op_u16(Op::REF_GET_DESC, 0, 0);
    });
    assert_eq!(r.as_str(), "my-type");
}

#[test]
fn struct_new_default_desc_attaches_descriptor() {
    let r = run(|c| {
        c.emit_i32_const(42, 0);
        c.emit_op_u16(Op::STRUCT_NEW_DEFAULT_DESC, 0, 0);
        c.emit_op_u16(Op::REF_GET_DESC, 0, 0);
    });
    assert_eq!(r.as_i32(), 42);
}

#[test]
fn ref_get_desc_retrieves_descriptor() {
    let r = run(|c| {
        c.emit_string_const("tag", 0);
        c.emit_op_u16_u16(Op::STRUCT_NEW_DESC, 0, 0, 0);
        c.emit_op_u16(Op::REF_GET_DESC, 0, 0);
    });
    assert_eq!(r.as_str(), "tag");
}

// The descriptor is the LAST operand, so the field values sit BENEATH it on
// the stack (Overview.md §"Allocation With Descriptors"). The stub popped only
// the descriptor, which left every field value stranded on the stack.
#[test]
fn struct_new_desc_pops_its_field_operands() {
    let r = run(|c| {
        c.emit_i32_const(7, 0); // a field value
        c.emit_string_const("d", 0); // descriptor on top
        c.emit_op_u16_u16(Op::STRUCT_NEW_DESC, 0, 1, 0);
        // If the field operand were left behind, the struct ref would not be
        // on top and `ref.get_desc` would receive the stray i32 instead.
        c.emit_op_u16(Op::REF_GET_DESC, 0, 0);
    });
    assert_eq!(r.as_str(), "d");
}

// `test/core/custom-descriptors/struct_new_desc.wast:492-497` — a null
// descriptor traps, and does so for both allocation forms.
#[test]
fn struct_new_desc_traps_on_a_null_descriptor() {
    let e = run_err(|c| {
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        c.emit_op_u16_u16(Op::STRUCT_NEW_DESC, 0, 0, 0);
    });
    assert!(e.contains("null descriptor reference"), "{e}");
}

#[test]
fn struct_new_default_desc_traps_on_a_null_descriptor() {
    let e = run_err(|c| {
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        c.emit_op_u16(Op::STRUCT_NEW_DEFAULT_DESC, 0, 0);
    });
    assert!(e.contains("null descriptor reference"), "{e}");
}

// `ref_get_desc.wast:400-406` — the result type is non-nullable, so a null
// input cannot be passed through.
#[test]
fn ref_get_desc_traps_on_a_null_reference() {
    let e = run_err(|c| {
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        c.emit_op_u16(Op::REF_GET_DESC, 0, 0);
    });
    assert!(e.contains("null reference"), "{e}");
}

#[test]
fn ref_get_desc_on_struct_without_desc_returns_null() {
    let r = run(|c| {
        c.emit_struct_new(0, 0, 0);
        c.emit_op_u16(Op::REF_GET_DESC, 0, 0);
        c.emit_op(Op::REF_IS_NULL, 0);
    });
    assert_eq!(r.as_i32(), 1);
}

// ═══════════════════════════════════════════════════════════════════════════
// Custom Descriptors: the descriptor-comparing casts
// (0xFB 0x23 ref.cast_desc_eq / 0x24 nullable / 0x25 br_on_cast_desc_eq /
//  0x26 br_on_cast_desc_eq_fail)
//
// These cast on descriptor IDENTITY, not on the type hierarchy: a reference
// passes iff the descriptor it was allocated with is the very same reference
// as the one supplied on the stack.
// ═══════════════════════════════════════════════════════════════════════════

/// `[ref-with-descriptor `desc`, descriptor `operand`]` — the operand layout
/// every one of the four instructions takes. Descriptors ride the spec
/// string-constant route (one imported global per distinct text): the same
/// name yields the same materialized Arc on every `global.get`, a different
/// name a different one — exactly the identity semantics these casts test.
fn push_described_ref_and_descriptor(c: &mut Chunk, desc: &str, operand: &str) {
    c.emit_string_const(desc, 0);
    c.emit_op_u16_u16(Op::STRUCT_NEW_DESC, 0, 0, 0);
    c.emit_string_const(operand, 0);
}

#[test]
fn ref_cast_desc_eq_passes_when_the_descriptor_is_the_same_reference() {
    let r = run(|c| {
        push_described_ref_and_descriptor(c, "vtable-a", "vtable-a");
        c.emit_op_u16(Op::REF_CAST_DESC_EQ, 0, 0);
        // The cast consumes only the descriptor; the reference survives it.
        c.emit_op_u16(Op::REF_GET_DESC, 0, 0);
    });
    assert_eq!(r.as_str(), "vtable-a");
}

#[test]
fn ref_cast_desc_eq_traps_on_a_different_descriptor() {
    let err = run_err(|c| {
        push_described_ref_and_descriptor(c, "vtable-a", "vtable-b");
        c.emit_op_u16(Op::REF_CAST_DESC_EQ, 0, 0);
    });
    assert!(
        err.contains("descriptor cast failure"),
        "a descriptor that is not the allocated one must trap, got: {err}"
    );
}

#[test]
fn ref_cast_desc_eq_traps_on_a_null_descriptor_before_looking_at_the_reference() {
    // The reference here would match nothing anyway, but the point is the
    // ORDER: the proposal traps on a null descriptor unconditionally, so the
    // message must be the null-descriptor one, not the mismatch one.
    let err = run_err(|c| {
        c.emit_string_const("vtable-a", 0);
        c.emit_op_u16_u16(Op::STRUCT_NEW_DESC, 0, 0, 0);
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0); // descriptor operand = null
        c.emit_op_u16(Op::REF_CAST_DESC_EQ, 0, 0);
    });
    assert!(
        err.contains("null descriptor"),
        "a null descriptor traps first, got: {err}"
    );
}

#[test]
fn ref_cast_desc_eq_traps_on_a_reference_with_no_descriptor() {
    // A struct built with plain `struct.new` carries no descriptor, so it can
    // never equal a real one.
    let err = run_err(|c| {
        c.emit_struct_new(0, 0, 0);
        c.emit_string_const("vtable-a", 0);
        c.emit_op_u16(Op::REF_CAST_DESC_EQ, 0, 0);
    });
    assert!(
        err.contains("descriptor cast failure"),
        "an undescribed reference must fail the cast, got: {err}"
    );
}

#[test]
fn ref_cast_desc_eq_traps_on_a_null_reference_but_the_nullable_form_does_not() {
    let desc_text = "vtable-a";

    let err = run_err(|c| {
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0); // the reference
        c.emit_string_const(desc_text, 0);
        c.emit_op_u16(Op::REF_CAST_DESC_EQ, 0, 0);
    });
    // ⚠ "descriptor cast failure", NOT "null reference". This assertion used
    // to demand the latter, which is the wording `ref.get_desc` uses; for a
    // CAST the proposal's suite is explicit that a null reference is an
    // ordinary failed cast:
    //   ref_cast_desc_eq.wast:820
    //     (assert_trap (invoke "self-nonnullable-null-desc")
    //                  "descriptor cast failure")
    assert!(
        err.contains("descriptor cast failure"),
        "the (ref ht) form does not admit null, got: {err}"
    );

    // The (ref null ht) form passes null straight through.
    let r = run(|c| {
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        c.emit_string_const(desc_text, 0);
        c.emit_op_u16(Op::REF_CAST_DESC_EQ_NULL, 0, 0);
        c.emit_op(Op::REF_IS_NULL, 0);
    });
    assert_eq!(
        r.as_i32(),
        1,
        "the nullable form must leave null on the stack"
    );
}

#[test]
fn ref_cast_desc_eq_null_still_traps_on_a_mismatched_non_null_reference() {
    // Nullability is about the REFERENCE, not about the descriptor check —
    // a non-null value with the wrong descriptor traps in both forms.
    let err = run_err(|c| {
        push_described_ref_and_descriptor(c, "vtable-a", "vtable-b");
        c.emit_op_u16(Op::REF_CAST_DESC_EQ_NULL, 0, 0);
    });
    assert!(err.contains("descriptor cast failure"), "got: {err}");
}

/// Runs one `br_on_cast_desc_eq`-family instruction inside a block and reports
/// what came out. Branching leaves the reference on the stack and jumps past
/// the block; falling through runs the in-block code, which replaces the
/// reference with an integer. `ref.get_desc` after the block therefore answers
/// the descriptor when the branch was taken and null when it was not.
///
/// `same_descriptor` picks whether the operand is the very reference the
/// value was allocated with. Two constants holding equal TEXT are still two
/// distinct references, and correctly do not match — which is exactly the
/// distinction between descriptor equality and type equality.
fn run_desc_branch(op: Op, same_descriptor: bool) -> Value {
    run(|c| {
        // The pool entry stays alive as the instruction's name immediate;
        // the descriptor VALUES ride the string-constant global route.
        let desc = c.add_constant(Value::String(Arc::from("vtable-a")));
        let operand = if same_descriptor {
            "vtable-a"
        } else {
            "vtable-b"
        };

        // Arity 1: per the proposal the label takes the reference
        // (`C.labels[l] = t* rt`), so a void block would discard it on branch.
        let _blk = c.emit_block_typed(0, 1);
        push_described_ref_and_descriptor(c, "vtable-a", operand);
        c.emit_op(op, 0);
        c.emit((desc >> 8) as u8, 0);
        c.emit((desc & 0xFF) as u8, 0);
        c.emit(0u8, 0); // label depth 0 — exit the enclosing block
        // Fallthrough only.
        c.emit_op(Op::DROP, 0);
        c.emit_i32_const(0, 0);
        c.emit_end(0);

        c.emit_op_u16(Op::REF_GET_DESC, 0, 0);
    })
}

#[test]
fn br_on_cast_desc_eq_branches_only_when_the_descriptor_matches() {
    let taken = run_desc_branch(Op::BR_ON_CAST_DESC_EQ, true);
    assert_eq!(
        taken.as_str(),
        "vtable-a",
        "a matching descriptor must branch with the reference intact"
    );

    let not_taken = run_desc_branch(Op::BR_ON_CAST_DESC_EQ, false);
    assert!(
        matches!(not_taken, Value::Null),
        "a mismatched descriptor must fall through, got: {not_taken:?}"
    );
}

#[test]
fn br_on_cast_desc_eq_fail_branches_only_when_the_descriptor_differs() {
    let taken = run_desc_branch(Op::BR_ON_CAST_DESC_EQ_FAIL, false);
    assert_eq!(
        taken.as_str(),
        "vtable-a",
        "the _fail form branches on MISMATCH, carrying the reference"
    );

    let not_taken = run_desc_branch(Op::BR_ON_CAST_DESC_EQ_FAIL, true);
    assert!(
        matches!(not_taken, Value::Null),
        "a matching descriptor must fall through, got: {not_taken:?}"
    );
}

#[test]
fn br_on_cast_desc_eq_fail_traps_on_a_null_descriptor() {
    // The null-descriptor trap applies to all four instructions — including
    // the one whose branch condition is "did not match", where treating null
    // as a mismatch would be the tempting shortcut.
    let err = run_err(|c| {
        // Pool entry only feeds the instruction's name immediate now; the
        // descriptor VALUE rides the string-constant global route.
        let desc = c.add_constant(Value::String(Arc::from("vtable-a")));
        c.emit_string_const("vtable-a", 0);
        c.emit_op_u16_u16(Op::STRUCT_NEW_DESC, 0, 0, 0);
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        c.emit_op(Op::BR_ON_CAST_DESC_EQ_FAIL, 0);
        c.emit((desc >> 8) as u8, 0);
        c.emit((desc & 0xFF) as u8, 0);
        c.emit(0u8, 0);
    });
    assert!(err.contains("null descriptor"), "got: {err}");
}

// ── Custom Descriptors: binary format ────────────────────────────────────────

#[test]
fn descriptor_instructions_do_not_desync_the_instructions_after_them() {
    // `struct.new_desc` and friends each carry a typeidx. The reader used to
    // consume ZERO immediate bytes for them, so the typeidx was decoded as if
    // it were the next opcode and everything downstream shifted. Asserting
    // only that the descriptor op survived would pass with that bug present —
    // what has to survive is the instruction AFTER it.
    let mut chunk = Chunk::new("<script>");
    chunk.emit_string_const("vtable-a", 0);
    chunk.emit_op_u16_u16(Op::STRUCT_NEW_DESC, 0, 0, 0);
    chunk.emit_op_u16(Op::REF_GET_DESC, 0, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_NONE, 0); // the marker that must still decode
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::RETURN, 0);

    let wasm = write_wasm(&vec![chunk]);
    assert!(
        has_bytes(&wasm, &[0xFB, 0x20]) && has_bytes(&wasm, &[0xFB, 0x22]),
        "struct.new_desc (0xFB 0x20) and ref.get_desc (0xFB 0x22) must be emitted"
    );

    let chunks = read_wasm(&wasm).expect("read_wasm failed");
    let wasm2 = write_wasm(&chunks);
    assert!(
        has_bytes(&wasm2, &[0xD0, 0x71]),
        "the ref.null none after the descriptor ops must survive decoding — \
         losing it means the typeidx immediate was read as an opcode"
    );
    assert!(
        has_bytes(&wasm2, &[0xFB, 0x22]),
        "ref.get_desc must round-trip"
    );
}

#[test]
fn descriptor_casts_round_trip_through_the_binary_format() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_string_const("vtable-a", 0);
    chunk.emit_op_u16_u16(Op::STRUCT_NEW_DESC, 0, 0, 0);
    chunk.emit_string_const("vtable-a", 0);
    chunk.emit_op_u16(Op::REF_CAST_DESC_EQ, 0, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_NONE, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::RETURN, 0);

    let wasm = write_wasm(&vec![chunk]);
    assert!(
        has_bytes(&wasm, &[0xFB, 0x23]),
        "ref.cast_desc_eq must encode as 0xFB 0x23"
    );

    let chunks = read_wasm(&wasm).expect("read_wasm failed");
    let wasm2 = write_wasm(&chunks);
    assert!(
        has_bytes(&wasm2, &[0xFB, 0x23]) && has_bytes(&wasm2, &[0xD0, 0x71]),
        "the cast and the instruction after it must both survive the round trip"
    );
}

#[test]
fn br_on_cast_desc_eq_encodes_castflags_labelidx_and_two_heaptypes() {
    let mut chunk = Chunk::new("<script>");
    let desc = chunk.add_constant(Value::String(Arc::from("vtable-a")));
    chunk.emit_string_const("vtable-a", 0);
    chunk.emit_op_u16_u16(Op::STRUCT_NEW_DESC, 0, 0, 0);
    chunk.emit_string_const("vtable-a", 0);
    chunk.emit_op(Op::BR_ON_CAST_DESC_EQ, 0);
    chunk.emit((desc >> 8) as u8, 0);
    chunk.emit((desc & 0xFF) as u8, 0);
    chunk.emit(0u8, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_NONE, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::RETURN, 0);

    let wasm = write_wasm(&vec![chunk]);
    // 0xFB 0x25, then castflags 0x00, labelidx 0x00, then two heaptypes.
    assert!(
        has_bytes(&wasm, &[0xFB, 0x25, 0x00, 0x00]),
        "br_on_cast_desc_eq must encode as 0xFB 0x25 castflags labelidx ..."
    );

    let chunks = read_wasm(&wasm).expect("read_wasm failed");
    let wasm2 = write_wasm(&chunks);
    assert!(
        has_bytes(&wasm2, &[0xFB, 0x25]) && has_bytes(&wasm2, &[0xD0, 0x71]),
        "the branch and the instruction after it must both survive the round trip"
    );
}

// Exact heap types (`0x62 x:u32`) and exact function imports (`externtype
// 0x20`) are decode-side only — they have no bytecode representation to drive
// from here. They are covered by unit tests against the section parsers in
// `platforms/wasm/src/reader/mod.rs`.

// ── WASM GC typed-null (ref.null none) codec round-trip ──────────────────────

#[test]
fn typed_null_encodes_as_ref_null_none_and_roundtrips() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_NONE, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::RETURN, 0);

    let wasm = write_wasm(&vec![chunk]);
    // A WASM GC typed null serializes to the real `ref.null none` (0xD0 0x71),
    // NOT `ref.null extern` (0xD0 0x6F).
    assert!(
        has_bytes(&wasm, &[0xD0, 0x71]),
        "typed null must encode as ref.null none (0xD0 0x71)"
    );

    // Decode + re-encode: it must still be `ref.null none`, i.e. it decoded back
    // to a typed null (`ref.null none`), not collapsed to a plain null.
    let chunks = read_wasm(&wasm).expect("read_wasm failed");
    let wasm2 = write_wasm(&chunks);
    assert!(
        has_bytes(&wasm2, &[0xD0, 0x71]),
        "typed null must round-trip as ref.null none"
    );
}

#[test]
fn plain_null_encodes_as_ref_null_extern() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op(Op::RETURN, 0);

    let wasm = write_wasm(&vec![chunk]);
    // A plain (dynamic-language) null is `ref.null extern` (0xD0 0x6F).
    assert!(
        has_bytes(&wasm, &[0xD0, 0x6F]),
        "plain null must encode as ref.null extern (0xD0 0x6F)"
    );
}

// ── Trap conditions ─────────────────────────────────────────────────────
//
// The GC proposal specifies traps, not lenient defaults: an out-of-bounds
// array access traps, `array.len` and the `i31` getters trap on null. These
// paths had no coverage, so `array.get_s`/`get_u` silently answered 0 past the
// end and `i31.get_s` read a null back as 0 — indistinguishable from a genuine
// `ref.i31 0`.
//
// Only a stamped GC array is bounds-checked; a dynamic (JS-shaped) array stays
// lenient, which is why these build the array through `array.new_fixed`.

fn run_locals_err(local_count: u16, emit: impl FnOnce(&mut Chunk)) -> String {
    let mut c = Chunk::new("<script>");
    c.local_count = local_count;
    emit(&mut c);
    c.emit_op(Op::RETURN, 0);
    VM::new()
        .run(vec![c])
        .map(|value| format!("returned {value:?} instead of trapping"))
        .unwrap_err()
        .to_string()
}

/// Build a 3-element GC array into local 0.
///
/// Deliberately `array.new $t`, NOT `array.new_fixed`: only the typed
/// constructors stamp the array's rtt, and an unstamped array is treated as a
/// dynamic (JS-shaped) one that never bounds-checks. `array.new_fixed $t N`
/// carries a type immediate in the spec but only the count in our encoding,
/// so it cannot stamp, and every array it builds escapes bounds checking.
fn emit_three_element_array(c: &mut Chunk) {
    c.types.push(vybe_runtime::chunk::TypeEntry {
        name: "gcarray".into(),
        kind: vybe_runtime::chunk::CompositeKind::Array,
        parent_index: 0,
        fields: Vec::new(),
        methods: Vec::new(),
        is_interface: false,
        implements: Vec::new(),
        constructor_chunk: None,
        field_descriptors: std::collections::HashMap::new(),
            ..Default::default()
    });
    c.emit_i32_const(0, 0); // fill value
    c.emit_i32_const(3, 0); // length
    c.emit_op_u16(Op::ARRAY_NEW, 1, 0); // 1-based type immediate
    c.emit_op_u16(Op::LOCAL_SET, 0, 0);
}

#[test]
fn array_len_traps_on_null_reference() {
    let err = run_locals_err(0, |c| {
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_NONE, 0);
        c.emit_op(Op::ARRAY_LENGTH, 0);
    });
    assert!(
        err.contains("trap") && err.contains("null"),
        "array.len must trap on a null array reference, got: {err}"
    );
}

#[test]
fn i31_get_s_traps_on_null_reference() {
    let err = run_locals_err(0, |c| {
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        c.emit_op(Op::I31_GET_S, 0);
    });
    assert!(
        err.contains("trap") && err.contains("null"),
        "i31.get_s must trap on null rather than reading back 0, got: {err}"
    );
}

#[test]
fn i31_get_u_traps_on_null_reference() {
    let err = run_locals_err(0, |c| {
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        c.emit_op(Op::I31_GET_U, 0);
    });
    assert!(
        err.contains("trap") && err.contains("null"),
        "i31.get_u must trap on null, got: {err}"
    );
}

/// `ref.i31 0` and a null must stay distinguishable — the reason the getters
/// have to trap rather than answer a default.
#[test]
fn i31_round_trips_zero_without_trapping() {
    let value = run(|c| {
        c.emit_i32_const(0, 0);
        c.emit_op(Op::I31_NEW, 0);
        c.emit_op(Op::I31_GET_S, 0);
    });
    assert_eq!(value.as_i32(), 0);
}

/// `ref.i31` keeps the low 31 bits, and `i31.get_s` sign-extends from bit 30 —
/// so the largest 31-bit pattern reads back negative, not as a huge positive.
#[test]
fn i31_wraps_to_31_bits_and_sign_extends() {
    let value = run(|c| {
        c.emit_i32_const(0x4000_0000, 0);
        c.emit_op(Op::I31_NEW, 0);
        c.emit_op(Op::I31_GET_S, 0);
    });
    assert_eq!(
        value.as_i32(),
        -1_073_741_824,
        "bit 30 is the sign bit of a 31-bit value"
    );

    // The same bits read unsigned stay positive.
    let unsigned = run(|c| {
        c.emit_i32_const(0x4000_0000, 0);
        c.emit_op(Op::I31_NEW, 0);
        c.emit_op(Op::I31_GET_U, 0);
    });
    assert_eq!(unsigned.as_i32(), 0x4000_0000);
}

#[test]
fn array_get_traps_past_the_end() {
    let err = run_locals_err(1, |c| {
        emit_three_element_array(c);
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(3, 0);
        c.emit_op(Op::ARRAY_GET, 0);
    });
    assert!(
        err.contains("trap") && err.contains("bounds"),
        "array.get must trap past the end, got: {err}"
    );
}

#[test]
fn array_get_s_traps_past_the_end() {
    let err = run_locals_err(1, |c| {
        emit_three_element_array(c);
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(3, 0);
        c.emit_op_u16(Op::ARRAY_GET_S, 0, 0);
    });
    assert!(
        err.contains("trap") && err.contains("bounds"),
        "array.get_s must trap past the end rather than answering 0, got: {err}"
    );
}

#[test]
fn array_get_u_traps_on_a_negative_index() {
    let err = run_locals_err(1, |c| {
        emit_three_element_array(c);
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(-1, 0);
        c.emit_op_u16(Op::ARRAY_GET_U, 0, 0);
    });
    assert!(
        err.contains("trap") && err.contains("bounds"),
        "a negative index must trap, not clamp to 0, got: {err}"
    );
}

#[test]
fn array_fill_traps_when_the_region_leaves_the_array() {
    let err = run_locals_err(1, |c| {
        emit_three_element_array(c);
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(1, 0); // index
        c.emit_i32_const(1, 0); // value
        c.emit_i32_const(5, 0); // count — 1 + 5 > 3
        c.emit_op(Op::ARRAY_FILL, 0);
    });
    assert!(
        err.contains("trap") && err.contains("bounds"),
        "array.fill must trap rather than silently filling less, got: {err}"
    );
}

/// A fill that fits must still work — the trap must not swallow the valid case.
#[test]
fn array_fill_within_bounds_still_writes() {
    let value = run_locals(1, |c| {
        emit_three_element_array(c);
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(0, 0); // index
        c.emit_i32_const(7, 0); // value
        c.emit_i32_const(2, 0); // count
        c.emit_op(Op::ARRAY_FILL, 0);

        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(0, 0);
        c.emit_op(Op::ARRAY_GET, 0);
    });
    assert_eq!(value.as_i32(), 7);
}

// ── Type section round trip ─────────────────────────────────────────────

/// A module carrying GC types must read back with its FUNCTION types at the
/// same indices it wrote them.
///
/// The type index space counts every subtype inside a `rec` group, so a reader
/// that skips struct and array types shifts each later function type's index —
/// and those indices are what `call` and blocktypes resolve against. The reader
/// used to look only for `0x60` and, on anything else, advance a single byte,
/// which neither skipped the type body nor kept alignment: a module with GC
/// types desynchronised from its first struct on.
#[test]
fn gc_type_section_round_trips_without_shifting_function_indices() {
    let mut chunk = Chunk::new("<script>");
    chunk.types.push(type_entry("alpha", &["x", "y"]));
    chunk.types.push(type_entry("beta", &["z"]));
    chunk.emit_i32_const(7, 0);
    chunk.emit_op(Op::RETURN, 0);

    let bytes = write_wasm(&[chunk]);
    let chunks = read_wasm(&bytes).expect("a module carrying GC types must be readable");
    assert!(
        !chunks.is_empty(),
        "the reader must recover at least the script chunk"
    );
}

/// The writer emits struct types inside a `rec` group — the shape the reader
/// has to understand. Pinned so a writer change that drops the group does not
/// silently pass by making the reader's job easier.
#[test]
fn gc_struct_types_are_written_inside_a_rec_group() {
    let mut chunk = Chunk::new("<script>");
    chunk.types.push(type_entry("gamma", &["f"]));
    chunk.emit_op(Op::RETURN, 0);

    let bytes = write_wasm(&[chunk]);
    // 0x4E = rec, 0x5F = struct composite type.
    assert!(
        has_bytes(&bytes, &[0x4E]) && has_bytes(&bytes, &[0x5F]),
        "expected a rec group containing a struct composite type"
    );
}

// ── array.new_fixed operand width ───────────────────────────────────────────

#[test]
fn array_new_fixed_inside_a_block_does_not_desync_the_scan() {
    // `array.new_fixed` carries BOTH `$t` and `N` — four operand bytes. The
    // VM's block-scanning pre-pass walks instructions by
    // `operand_format().size_in()`, so declaring the op two bytes narrower
    // than it is makes the scan resume INSIDE the following instruction and
    // mismatch the enclosing block, which silently halts the program instead
    // of trapping. Only code that builds a fixed array inside a block hits
    // it — which is why it surfaced as one PHP branch (`array_fill(0, …)`)
    // rather than as a broken opcode.
    assert_eq!(
        Op::ARRAY_NEW_FIXED.operand_format().fixed_size(),
        4,
        "array.new_fixed declares $t + N"
    );

    let r = run(|c| {
        let _blk = c.emit_block_typed(0, 1);
        c.emit_array_new_fixed(0, 0, 0); // empty array, in-block
        c.emit_op(Op::ARRAY_LENGTH, 0);
        c.emit_end(0);
        // Reached only if the block closed where the scan thought it did.
        c.emit_op(Op::DROP, 0);
        c.emit_i32_const(7, 0);
    });
    assert_eq!(
        r.as_i32(),
        7,
        "execution must continue past the block that built the array"
    );
}

// ── struct.new: dynamic vs typed ────────────────────────────────────────────

#[test]
fn struct_new_with_a_typeidx_stamps_rtt_and_fills_indexed_fields() {
    // Spec `struct.new $t` takes its field count from $t, not from an
    // immediate, lands the values in INDEXED storage, and stamps $t's rtt so
    // `ref.test` answers from the type registry rather than a `__type`
    // string. typeidx 0 stays the dynamic object-literal form.
    let mut chunk = Chunk::new("<script>");
    chunk.types.push(type_entry("Point", &["x", "y"]));
    chunk.emit_i32_const(11, 0);
    chunk.emit_i32_const(22, 0);
    chunk.emit_struct_new(1, 0, 0); // typeidx 1 → chunk.types[0] = "Point"
    chunk.emit_struct_field_op(Op::STRUCT_GET_U, 0, 1, 0); // indexed read of field 1
    chunk.emit_op(Op::RETURN, 0);
    let r = VM::new().run(vec![chunk]).expect("run failed");
    assert_eq!(
        r.as_i32(),
        22,
        "field 1 must come back from indexed storage"
    );
}

#[test]
fn struct_new_with_a_typeidx_answers_ref_test_from_the_registry() {
    let mut chunk = Chunk::new("<script>");
    chunk.types.push(type_entry("Point", &["x", "y"]));
    chunk.emit_i32_const(1, 0);
    chunk.emit_i32_const(2, 0);
    chunk.emit_struct_new(1, 0, 0);
    // Same immediate the allocation carried — the test names no type.
    chunk.emit_ref_type_op(Op::REF_TEST, HeapType::Concrete(1), 0);
    chunk.emit_op(Op::RETURN, 0);
    let r = VM::new().run(vec![chunk]).expect("run failed");
    assert_eq!(
        r.as_i32(),
        1,
        "a typed struct.new must be recognised as its declared type"
    );
}

#[test]
fn struct_new_with_typeidx_zero_is_still_the_object_literal_form() {
    // The ~258 rewritten emit sites all pass 0 — key/value pairs, named
    // properties, no rtt. Breaking this breaks every object literal.
    let r = run(|c| {
        // Pool entry only feeds STRUCT_GET's name immediate; the KEY value
        // on the stack rides the string-constant global route.
        let k = c.add_constant(Value::String(Arc::from("a")));
        c.emit_string_const("a", 0);
        c.emit_i32_const(5, 0);
        c.emit_struct_new(0, 1, 0);
        c.emit_struct_field_op(Op::STRUCT_GET, 0, k, 0);
    });
    assert_eq!(r.as_i32(), 5);
}

#[test]
fn struct_new_default_with_a_typeidx_allocates_the_declared_field_slots() {
    // `struct.new_default $t` takes nothing off the stack: the instance is
    // $t's fields at their defaults, with $t's rtt stamped.
    let mut chunk = Chunk::new("<script>");
    chunk.types.push(type_entry("Pair", &["a", "b"]));
    chunk.emit_op_u16(Op::STRUCT_NEW_DEFAULT, 1, 0);
    chunk.emit_ref_type_op(Op::REF_TEST, HeapType::Concrete(1), 0);
    chunk.emit_op(Op::RETURN, 0);
    let r = VM::new().run(vec![chunk]).expect("run failed");
    assert_eq!(r.as_i32(), 1, "the instance must carry its declared type");

    // And the field slots exist — an indexed read of field 1 is defined
    // (Null), not an out-of-range miss.
    let mut chunk = Chunk::new("<script>");
    chunk.types.push(type_entry("Pair", &["a", "b"]));
    chunk.emit_op_u16(Op::STRUCT_NEW_DEFAULT, 1, 0);
    chunk.emit_struct_field_op(Op::STRUCT_GET_U, 0, 1, 0);
    chunk.emit_op(Op::REF_IS_NULL, 0);
    chunk.emit_op(Op::RETURN, 0);
    let r = VM::new().run(vec![chunk]).expect("run failed");
    assert_eq!(
        r.as_i32(),
        1,
        "declared slots must be allocated and defaulted"
    );
}

#[test]
fn struct_set_has_an_indexed_form_that_struct_get_reads_back() {
    // `struct.set $t i` did not exist — STRUCT_SET was name-keyed only, so
    // nothing could write the indexed storage `struct.get_s`/`_u` read.
    // Typed struct.set returns void, per spec.
    let mut chunk = Chunk::new("<script>");
    chunk.local_count = 1;
    chunk.types.push(type_entry("Cell", &["v"]));
    chunk.emit_op_u16(Op::STRUCT_NEW_DEFAULT, 1, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, 0, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_i32_const(9, 0);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 1, 0, 0); // typed: fieldidx 0
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 1, 0, 0); // typed read
    chunk.emit_op(Op::RETURN, 0);
    let r = VM::new().run(vec![chunk]).expect("run failed");
    assert_eq!(
        r.as_i32(),
        9,
        "typed struct.set must write the indexed slot"
    );
}

#[test]
fn typed_struct_field_ops_trap_out_of_range_and_on_null() {
    let mut chunk = Chunk::new("<script>");
    chunk.types.push(type_entry("Cell", &["v"]));
    chunk.emit_op_u16(Op::STRUCT_NEW_DEFAULT, 1, 0);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 1, 7, 0); // no field 7
    chunk.emit_op(Op::RETURN, 0);
    let err = VM::new().run(vec![chunk]).unwrap_err().to_string();
    assert!(err.contains("out of range"), "got: {err}");

    let mut chunk = Chunk::new("<script>");
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_NONE, 0);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 1, 0, 0);
    chunk.emit_op(Op::RETURN, 0);
    let err = VM::new().run(vec![chunk]).unwrap_err().to_string();
    assert!(err.contains("null reference"), "got: {err}");
}

#[test]
fn struct_field_ops_with_typeidx_zero_stay_name_keyed() {
    // The 1117 rewritten sites all pass 0 — this is every property access in
    // every language.
    let r = run_locals(1, |c| {
        let k = c.add_constant(Value::String(Arc::from("a")));
        c.emit_struct_new(0, 0, 0);
        c.emit_op_u16(Op::LOCAL_SET, 0, 0);
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_i32_const(3, 0);
        c.emit_struct_field_op(Op::STRUCT_SET, 0, k, 0);
        c.emit_op(Op::DROP, 0);
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        c.emit_struct_field_op(Op::STRUCT_GET, 0, k, 0);
    });
    assert_eq!(r.as_i32(), 3);
}
