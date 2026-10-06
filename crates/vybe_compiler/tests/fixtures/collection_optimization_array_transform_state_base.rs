//! Unit proofs of the actual crate-private inline seam; no substitute dispatcher.
use super::*;
use crate::primitives::{functions, globals, platforms};
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::{Value, VM};

fn array(values: Vec<Value>) -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new_array(values)))
}
fn ints(values: &[i32]) -> Value {
    array(values.iter().copied().map(Value::I32).collect())
}
fn items(value: &Value) -> Vec<Value> {
    let Value::Object(object) = value else {
        panic!("array required")
    };
    let object = object.lock().unwrap();
    let ObjectKind::Array(values) = &object.kind else {
        panic!("array required")
    };
    values.clone()
}
fn assert_values(actual: &Value, expected: &Value) {
    if let (Value::Object(_), Value::Object(_)) = (actual, expected) {
        let (a, b) = (items(actual), items(expected));
        assert_eq!(a.len(), b.len());
        for (a, b) in a.iter().zip(&b) {
            assert_values(a, b);
        }
    } else {
        assert_eq!(actual, expected);
    }
}
fn inputs(name: &str) -> ([Value; 3], Value) {
    let mask = array(vec![
        Value::Bool(true),
        Value::Bool(false),
        Value::Bool(true),
        Value::Bool(false),
    ]);
    let column = name.contains("_column");
    if name.starts_with("__stdlib_array_pack_vector_") {
        (
            [ints(&[10, 20, 30, 40]), mask, ints(&[7, 8, 9, 6])],
            ints(&[10, 30, 9, 6]),
        )
    } else if name.starts_with("__stdlib_array_pack_") {
        (
            [ints(&[10, 20, 30, 40]), mask, Value::Null],
            ints(&[10, 30]),
        )
    } else if name.starts_with("__stdlib_array_unpack_") {
        (
            [ints(&[10, 30]), mask, Value::I32(-1)],
            ints(&[10, -1, 30, -1]),
        )
    } else if name.starts_with("__stdlib_array_merge_") {
        (
            [ints(&[10, 20, 30, 40]), ints(&[1, 2, 3, 4]), mask],
            ints(&[10, 2, 30, 4]),
        )
    } else if name.starts_with("__stdlib_array_reshape_pad_") {
        (
            [ints(&[1, 2]), ints(&[2, 2]), ints(&[8, 9])],
            if column {
                array(vec![ints(&[1, 8]), ints(&[2, 9])])
            } else {
                array(vec![ints(&[1, 2]), ints(&[8, 9])])
            },
        )
    } else {
        (
            [ints(&[1, 2, 3, 4]), ints(&[2, 2]), Value::Null],
            if column {
                array(vec![ints(&[1, 3]), ints(&[2, 4])])
            } else {
                array(vec![ints(&[1, 2]), ints(&[3, 4])])
            },
        )
    }
}
fn program(name: &str, inline: bool) -> Vec<Chunk> {
    let argc = shared_helper_arity(name).unwrap();
    let mut root = Chunk::new("<array-state-base>");
    root.local_count = 17;
    root.module_receiver_abi = vybe_runtime::chunk::ReceiverAbi::Parameter;
    // Deliberately opposite facts prove variant policy is temporary.
    root.module_instance_fields_are_own_properties = !name.contains("_own_fields");
    root.module_members_on_prototype = !name.contains("_prototype_members");
    let before = (
        root.module_instance_fields_are_own_properties,
        root.module_members_on_prototype,
    );
    root.emit_i32_const(99, 3);
    root.emit_op_u16(Op::LOCAL_SET, 0, 3);
    for (offset, key) in ["first", "second", "third"]
        .into_iter()
        .take(argc as usize)
        .enumerate()
    {
        globals::emit_read(&mut root, key, 3);
        root.emit_op_u16(Op::LOCAL_SET, 7 + offset as u16, 3);
    }
    let mut chunks = if inline {
        emit_inline_shared_helper(&mut root, name, 7, 4);
        assert_eq!(
            (
                root.module_instance_fields_are_own_properties,
                root.module_members_on_prototype
            ),
            before
        );
        vec![root]
    } else {
        functions::emit_ref_func(&mut root, 1, 0, 4);
        for arg in 0..argc as u16 {
            root.emit_op_u16(Op::LOCAL_GET, 7 + arg, 4);
        }
        crate::primitives::callable::emit_direct_invoke_chunk(&mut root, argc, 4);
        vec![root, build_shared_helper(name)]
    };
    // Return result with preserved caller state and evaluated input slot.
    chunks[0].emit_op_u16(Op::LOCAL_GET, 0, 5);
    chunks[0].emit_op_u16(Op::LOCAL_GET, 7, 5);
    chunks[0].emit_array_new_fixed(0, 3, 5);
    chunks[0].emit_op(Op::RETURN, 5);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
#[test]
fn all_48_variants_keep_fresh_input_state_bases_order_facts_and_native_results() {
    assert_eq!(SHARED_HELPERS.len(), 48);
    for &(name, _) in SHARED_HELPERS {
        let (inputs, expected) = inputs(name);
        for inline in [false, true] {
            let chunks = program(name, inline);
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
                .unwrap_or_else(|e| panic!("{name} inline={inline}: {e}"));
            let mut vm = VM::new();
            platforms::register_platforms_all(&mut vm);
            for (key, input) in ["first", "second", "third"].into_iter().zip(&inputs) {
                vm.set_global(key, input.clone());
            }
            let result = items(&vm.run(chunks).unwrap());
            assert_values(&result[0], &expected);
            assert_eq!(
                result[1],
                Value::I32(99),
                "{name} inline={inline} caller sentinel"
            );
            assert_eq!(
                result[2], inputs[0],
                "{name} inline={inline} input slot identity"
            );
        }
    }
}
#[test]
fn all_48_body_emissions_have_exact_shifted_locals_and_metadata_contract() {
    for &(name, _) in SHARED_HELPERS {
        let helper = build_shared_helper(name);
        let argc = shared_helper_arity(name).unwrap();
        assert_eq!(helper.arity, argc);
        assert_eq!(
            helper.module_instance_fields_are_own_properties,
            name.contains("_own_fields")
        );
        assert_eq!(
            helper.module_members_on_prototype,
            name.contains("_prototype_members")
        );
        let mut inline = Chunk::new("<shifted-array-body>");
        inline.local_count = 17;
        emit_inline_shared_helper(&mut inline, name, 7, 6);
        let prefix = 12 * argc as usize; // One GET+SET for every already-evaluated argument.
        let mut normalized = inline.code[prefix..].to_vec();
        let mut offset = 0;
        let mut boundaries = std::collections::HashSet::new();
        while offset < normalized.len() {
            boundaries.insert(offset);
            let op = Op::decode(
                u16::from_be_bytes(normalized[offset..offset + 2].try_into().unwrap()),
                u16::from_be_bytes(normalized[offset + 2..offset + 4].try_into().unwrap()),
            )
            .unwrap();
            let width = 4 + op.operand_format().size_in(&normalized, offset + 4);
            if matches!(op, Op::LOCAL_GET | Op::LOCAL_SET | Op::LOCAL_TEE) {
                let slot =
                    u16::from_be_bytes(normalized[offset + 4..offset + 6].try_into().unwrap());
                normalized[offset + 4..offset + 6]
                    .copy_from_slice(&(slot.checked_sub(17).unwrap()).to_be_bytes());
            }
            offset += width;
        }
        assert_eq!(
            normalized,
            &helper.code[..helper.code.len() - 4],
            "{name} exact body except fresh local base"
        );
        assert_eq!(
            inline.local_count,
            helper.local_count + 17,
            "{name} exact state/scratch reservation"
        );
        let expected: std::collections::HashSet<_> = helper
            .block_i32_results
            .iter()
            .map(|key| key + prefix)
            .collect();
        assert_eq!(
            inline.block_i32_results, expected,
            "{name} appended I32 block offsets"
        );
        for key in &helper.block_i32_results {
            assert!(boundaries.contains(key), "{name} metadata boundary");
            assert!(matches!(
                Op::decode(helper.read_u16(*key), helper.read_u16(*key + 2)),
                Some(Op::IF) | Some(Op::BLOCK)
            ));
            assert_eq!(
                &helper.code[key + 4..key + 6],
                &[0, 1],
                "{name} zero-parameter single-I32 block"
            );
        }
        assert!(helper.call_indirect_sigs.is_empty() && inline.call_indirect_sigs.is_empty());
        assert!(helper.call_indirect_canon.is_empty() && inline.call_indirect_canon.is_empty());
        assert!(helper.stack_switch_handlers.is_empty() && inline.stack_switch_handlers.is_empty());
        eprintln!(
            "array metadata {name}: I32blocks={} locals={}",
            helper.block_i32_results.len(),
            helper.local_count
        );
    }
    assert_eq!(shared_helper_arity("__stdlib_array_pack_row_made_up"), None);
}

#[test]
fn appended_metadata_keeps_existing_caller_markers_and_only_proven_body_blocks() {
    for &(name, _) in SHARED_HELPERS {
        let mut caller = Chunk::new("<array-metadata-prefix>");
        caller.local_count = 17;
        caller.emit_i32_const(1, 2);
        caller.emit_if_i32(2);
        caller.emit_i32_const(1, 2);
        caller.emit_else(2);
        caller.emit_i32_const(0, 2);
        caller.emit_end(2);
        caller.emit_op(Op::DROP, 2);
        let existing = caller.block_i32_results.clone();
        let original_len = caller.code.len();
        emit_inline_shared_helper(&mut caller, name, 7, 3);
        assert!(
            existing.is_subset(&caller.block_i32_results),
            "{name} prefix markers"
        );
        let helper = build_shared_helper(name);
        let body_start = original_len + 12 * shared_helper_arity(name).unwrap() as usize;
        let appended: std::collections::HashSet<_> = caller
            .block_i32_results
            .difference(&existing)
            .copied()
            .collect();
        assert_eq!(
            appended,
            helper
                .block_i32_results
                .iter()
                .map(|key| body_start + key)
                .collect(),
            "{name} appended metadata exact domain"
        );
        for key in appended {
            assert!(key >= original_len && key + 6 <= caller.code.len());
            assert!(matches!(
                Op::decode(caller.read_u16(key), caller.read_u16(key + 2)),
                Some(Op::IF) | Some(Op::BLOCK)
            ));
            assert_eq!(&caller.code[key + 4..key + 6], &[0, 1]);
        }
    }
}

#[test]
fn all_48_local_budgets_fit_at_boundary_and_reject_before_any_mutation() {
    for &(name, _) in SHARED_HELPERS {
        let budget = shared_inline_local_budget(name);
        let helper = build_shared_helper(name);
        assert_eq!(
            budget, helper.local_count,
            "{name} exact full state/scratch budget"
        );
        let largest_fitting = u16::MAX - budget;
        let mut fitting = Chunk::new("<largest-fitting-array-body>");
        fitting.local_count = largest_fitting;
        emit_inline_shared_helper(&mut fitting, name, 7, 8);
        assert_eq!(
            fitting.local_count,
            u16::MAX,
            "{name} largest fitting caller"
        );
        let argc = shared_helper_arity(name).unwrap() as u16;
        for (local_count, input_base) in [
            (largest_fitting + 1, 7), // Nested state exhausts representable count.
            (17, 17 - argc + 1),      // Inputs extend beyond allocated slots.
            (17, u16::MAX),           // Input end itself overflows its encoding.
        ] {
            let mut invalid = Chunk::new("<refused-array-body>");
            invalid.local_count = local_count;
            invalid.module_instance_fields_are_own_properties = true;
            invalid.module_members_on_prototype = true;
            invalid.emit_i32_const(1, 6);
            invalid.emit_if_i32(6);
            invalid.emit_i32_const(1, 6);
            invalid.emit_else(6);
            invalid.emit_i32_const(0, 6);
            invalid.emit_end(6);
            invalid.emit_op(Op::DROP, 6);
            let before = format!("{invalid:?}");
            let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                emit_inline_shared_helper(&mut invalid, name, input_base, 8)
            }));
            assert!(
                result.is_err(),
                "{name} must refuse locals={local_count} inputs={input_base}"
            );
            assert_eq!(
                format!("{invalid:?}"),
                before,
                "{name} refusal preserves all bytecode/constants/imports/local/metadata/facts"
            );
        }
    }
}
