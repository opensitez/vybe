//! Link runtime helper chunks into compiled output.
//!
//! Appends helper chunks to program. Emits a preamble in the script chunk
//! that creates function refs and stores them in globals (`__vybe_*`).
//!
//! Call sites do: `global_get "__vybe_sorted"` + `call_ref 1`
//!
//! On Vybe VM, `register_all` overwrites these globals with host fn objects.
//! On any other runtime, the globals hold the bundled helper function refs.
//! One binary, works everywhere, no unresolvable imports.

#[path = "generator_native_markers.rs"]
mod generator_native_markers;

use std::collections::{BTreeSet, HashMap};
use std::sync::OnceLock;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

// Private algorithm calls are compile-time references, not writable language
// globals. The existing u16 function operand is fixed up before execution;
// only an exact compiler-origin offset record authorizes that fixup.
const PRIVATE_HELPERS: &[(&str, u8)] = &[
    ("primitive.array_range_add", 2),
    ("primitive.array_range_lt", 2),
    ("__stdlib_erf", 1),
    ("__stdlib_random_next_unit", 0),
    ("__stdlib_levenshtein", 2),
    ("__stdlib_similar_text", 2),
    ("__stdlib_csv_parse_line", 3),
    ("__stdlib_csv_parse_document", 3),
    ("__stdlib_csv_format_minimal", 3),
    ("__stdlib_csv_format_whitespace", 3),
    ("primitive.reference_read", 1),
    ("primitive.byte_copy", 3),
    ("primitive.generator_yield_value", 1),
    ("primitive.generator_start_valid", 1),
    ("primitive.generator_start_current", 1),
    ("primitive.generator_start_value", 1),
];
const PRIVATE_REF_PREFIX: &str = "<private-helper-ref:";
const PRIVATE_BODY_PREFIX: &str = "<private-helper-body:";
const UNRESOLVED_PRIVATE_REF: u16 = u16::MAX;

fn private_helper_arity(export: &str) -> Option<u8> {
    PRIVATE_HELPERS
        .iter()
        .find_map(|&(name, arity)| (name == export).then_some(arity))
}

// Module-local algorithm bodies come directly from their primitive emitters.
// They are ordinary Wasm functions, with no polyfill compilation, prelude or
// writable language-global registration.
fn build_private_primitive(export: &str) -> Option<Chunk> {
    Some(match export {
        "primitive.array_range_add" | "primitive.array_range_lt" => crate::primitives::arrays::build_range_binary(export),
        "__stdlib_erf" => crate::primitives::math::build_erf(),
        "__stdlib_random_next_unit" => crate::primitives::random::build_next_unit(),
        "__stdlib_levenshtein" => crate::primitives::string_similarity::build_levenshtein(),
        "__stdlib_similar_text" => crate::primitives::string_similarity::build_similar_text(),
        "__stdlib_csv_parse_line" | "__stdlib_csv_parse_document"
        | "__stdlib_csv_format_minimal" | "__stdlib_csv_format_whitespace" => {
            crate::primitives::csv::build_shared_helper(export)
        }
        "primitive.reference_read" => crate::primitives::references::build_guarded_reference_read(),
        "primitive.byte_copy" => crate::primitives::pointers::build_byte_copy(),
        "primitive.generator_yield_value" => crate::primitives::generators::build_generator_yield_value(),
        "primitive.generator_start_valid" | "primitive.generator_start_current"
        | "primitive.generator_start_value" => {
            crate::primitives::generators::build_buffered_generator_start(export)
        }
        _ => return None,
    })
}

fn private_ref_marker(name: &str) -> Option<(&'static str, usize)> {
    let value = name.strip_prefix(PRIVATE_REF_PREFIX)?.strip_suffix('>')?;
    let (export, offset) = value.rsplit_once('@')?;
    let export = PRIVATE_HELPERS
        .iter()
        .find_map(|&(name, _)| (name == export).then_some(name))?;
    Some((export, offset.parse().ok()?))
}

fn private_body_marker(export: &str) -> String {
    format!("{PRIVATE_BODY_PREFIX}{export}>")
}

pub(crate) fn mark_private_helper_body(chunk: &mut Chunk) {
    use vybe_runtime::chunk::{LocalName, LocalOrigin};
    assert_eq!(private_helper_arity(&chunk.name), Some(chunk.arity));
    chunk.local_names.push(LocalName::new(
        0,
        private_body_marker(&chunk.name),
        LocalOrigin::Compiler,
    ));
}

fn private_helper_index(chunks: &[Chunk], export: &str) -> Option<usize> {
    use vybe_runtime::chunk::LocalOrigin;
    let arity = private_helper_arity(export)?;
    let marker = private_body_marker(export);
    chunks.iter().position(|chunk| {
        chunk.name == export
            && chunk.arity == arity
            && chunk.local_names.iter().any(|local| {
                local.slot == 0 && local.origin == LocalOrigin::Compiler && local.name == marker
            })
    })
}

/// Preserve only pending generated reference records while exporting a source
/// scope's debugger names. Source records stay intact and come last, so they
/// retain normal debugger precedence even when a scratch slot was reused.
pub(crate) fn finalize_source_local_names(
    chunk: &mut Chunk,
    names: &[vybe_runtime::chunk::LocalName],
) {
    let mut pending = pending_source_local_markers(chunk);
    pending.extend_from_slice(names);
    chunk.local_names = pending;
}

/// Transfer a completed scope's records instead of cloning every source name.
/// Call only after its last emission; upvalue naming reads Scope.locals.
pub(crate) fn finalize_source_local_names_owned(
    chunk: &mut Chunk,
    names: Vec<vybe_runtime::chunk::LocalName>,
) {
    let mut pending = pending_source_local_markers(chunk);
    if pending.is_empty() {
        chunk.local_names = names;
    } else {
        pending.extend(names);
        chunk.local_names = pending;
    }
}

fn pending_source_local_markers(chunk: &Chunk) -> Vec<vybe_runtime::chunk::LocalName> {
    use vybe_runtime::chunk::LocalOrigin;
    let mut seen = BTreeSet::new();
    chunk
        .local_names
        .iter()
        .filter(|local| {
            local.origin == LocalOrigin::Compiler
                && (private_ref_marker(&local.name).is_some()
                    || local.name.starts_with("<native-generator-next:")
                    || local.name.starts_with("<native-generator-resume:")
                    || local.name.starts_with("<native-generator-entry:")
                    || local.name.starts_with("<native-growable-array:"))
                && seen.insert((local.slot, local.name.clone()))
        })
        .cloned()
        .collect()
}

/// Outline one registered private algorithm without publishing a language
/// binding. Evaluated arguments are spilled in their original stack order.
pub(crate) fn emit_private_helper_call(chunk: &mut Chunk, export: &str, argc: u8, line: u32) {
    assert_eq!(private_helper_arity(export), Some(argc));
    // A zero-argument caller needs only a valid debugger record slot. Reuse
    // slot zero when available; otherwise reserve it once, without stack ops.
    let base = if argc == 0 {
        if chunk.local_count == 0 {
            chunk.alloc_scratch(1);
        }
        0
    } else {
        chunk.alloc_scratch(argc as u16)
    };
    for arg in (0..argc).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, base + arg as u16, line);
    }
    emit_private_helper_call_from_slots(chunk, export, base, argc, line);
}

/// Direct, uncaptured range-math calls have ordinary explicit Wasm signatures.
/// A private marker resolves the forward target at module finalization.
pub(crate) fn emit_private_range_call_from_slots(chunk: &mut Chunk, export: &str, base: u16, line: u32) {
    use vybe_runtime::chunk::{LocalName, LocalOrigin};
    assert!(matches!(export, "primitive.array_range_add" | "primitive.array_range_lt"));
    assert_eq!(private_helper_arity(export), Some(2));
    assert!((base as usize) + 2 <= chunk.local_count as usize);
    let mut declaration = Chunk::new(export);
    declaration.arity = 2;
    declaration.func_sig = Some(("externref,externref".into(), "externref".into()));
    for arg in 0..2 { chunk.emit_op_u16(Op::LOCAL_GET, base + arg, line); }
    let offset = chunk.code.len();
    chunk.emit_direct_defined_call(u32::MAX, &declaration, false, line).expect("uncaptured private range helper");
    chunk.local_names.push(LocalName::new(base, format!("{PRIVATE_REF_PREFIX}{export}@{offset}>"), LocalOrigin::Compiler));
}

fn private_range_direct(export: &str) -> bool {
    matches!(export, "primitive.array_range_add" | "primitive.array_range_lt")
}

/// Call a module-local primitive using values already retained in caller slots.
/// Avoid extra Object spills, which could extend observable reference lifetimes.
pub(crate) fn emit_private_helper_call_from_slots(
    chunk: &mut Chunk, export: &str, base: u16, argc: u8, line: u32,
) {
    use vybe_runtime::chunk::{LocalName, LocalOrigin};
    assert_eq!(private_helper_arity(export), Some(argc));
    assert!((base as usize) + (argc as usize) <= chunk.local_count as usize);
    assert!(base < chunk.local_count);
    // This one registered helper keeps strict String errors in the caller.
    // Evaluated arguments are already spilled exactly once above. Intrinsic
    // lengths are pure for validated Strings; do not expose a generic callback.
    if export == "__stdlib_similar_text" {
        for arg in 0..2 {
            chunk.emit_op_u16(Op::LOCAL_GET, base + arg, line);
            let length = chunk.add_import("wasm:js-string", "length");
            chunk.emit_call(length, 1, line);
            chunk.emit_op(Op::DROP, line);
        }
    }
    let offset = chunk.code.len();
    crate::primitives::functions::emit_ref_func(chunk, UNRESOLVED_PRIVATE_REF as usize, 0, line);
    chunk.local_names.push(LocalName::new(
        base,
        format!("{PRIVATE_REF_PREFIX}{export}@{offset}>"),
        LocalOrigin::Compiler,
    ));
    for arg in 0..argc {
        chunk.emit_op_u16(Op::LOCAL_GET, base + arg as u16, line);
    }
    crate::primitives::callable::emit_direct_invoke_chunk(chunk, argc, line);
}

fn instruction_boundaries(chunk: &Chunk) -> BTreeSet<usize> {
    let mut boundaries = BTreeSet::new();
    let mut offset = 0;
    while offset < chunk.code.len() {
        boundaries.insert(offset);
        assert!(
            offset + 4 <= chunk.code.len(),
            "truncated instruction before private fixup"
        );
        let op = Op::decode(chunk.read_u16(offset), chunk.read_u16(offset + 2))
            .expect("unknown instruction before private fixup");
        offset = offset
            .checked_add(4 + op.operand_format().size_in(&chunk.code, offset + 4))
            .expect("instruction size overflow before private fixup");
        assert!(
            offset <= chunk.code.len(),
            "truncated operand before private fixup"
        );
    }
    boundaries.insert(offset);
    boundaries
}

fn remap_offset(offset: usize, start: usize, end: usize, replacement_len: usize) -> usize {
    assert!(
        offset < start || offset >= end,
        "offset metadata overlaps a private call"
    );
    if offset < end {
        offset
    } else {
        offset
            .checked_sub(end - start)
            .and_then(|offset| offset.checked_add(replacement_len))
            .expect("offset metadata overflow during private inlining")
    }
}

fn inline_private_call(chunk: &mut Chunk, export: &str, offset: usize, slot: u16) {
    let argc = private_helper_arity(export).expect("registered private helper");
    let boundaries = instruction_boundaries(chunk);
    assert!(
        boundaries.contains(&offset),
        "private reference is not an instruction boundary"
    );
    let call = offset + 7 + argc as usize * 6;
    let end = call + 6;
    assert!(
        end <= chunk.code.len() && boundaries.contains(&end),
        "truncated private call sequence"
    );
    for arg in 0..argc {
        let load = offset + 7 + arg as usize * 6;
        assert_eq!(
            Op::decode(chunk.read_u16(load), chunk.read_u16(load + 2)),
            Some(Op::LOCAL_GET)
        );
        assert_eq!(chunk.read_u16(load + 4), slot + arg as u16);
    }
    assert_eq!(
        Op::decode(chunk.read_u16(call), chunk.read_u16(call + 2)),
        Some(Op::CALL_REF)
    );
    assert_eq!(
        &chunk.code[call + 4..end],
        &[argc, 1],
        "private call ABI was modified"
    );
    let line = chunk.lines[offset];
    let original_len = chunk.code.len();
    let original_i32_blocks = chunk.block_i32_results.clone();
    let original_private_refs: Vec<_> = chunk
        .local_names
        .iter()
        .filter(|local| private_ref_marker(&local.name).is_some())
        .map(|local| (local.slot, local.name.clone()))
        .collect();
    let metadata_sizes = (
        chunk.call_indirect_sigs.len(),
        chunk.call_indirect_canon.len(),
        chunk.stack_switch_handlers.len(),
        chunk.direct_defined_calls.len(),
    );
    // Append against the actual caller so constants/imports and scratch slots
    // remain in its existing index spaces, then move only the emitted bytes.
    match export {
        "primitive.byte_copy" => {
            for argument in 0..3 {
                chunk.emit_op_u16(Op::LOCAL_GET, slot + argument, line);
            }
            crate::primitives::pointers::emit_byte_copy_body(
                std::slice::from_mut(chunk), 0, line,
            );
        }
        "__stdlib_erf" => crate::primitives::math::emit_erf_body(chunk, slot, line),
        "__stdlib_similar_text" => crate::primitives::string_similarity::emit_similar_text_body(
            std::slice::from_mut(chunk), 0, slot, slot + 1, line,
        ),
        "__stdlib_levenshtein" => crate::primitives::string_similarity::emit_levenshtein_body(
            std::slice::from_mut(chunk),
            0,
            slot,
            slot + 1,
            line,
        ),
        "__stdlib_random_next_unit" => {
            crate::primitives::random::emit_next_unit_body(std::slice::from_mut(chunk), 0, line)
        }
        "__stdlib_csv_parse_line" => crate::primitives::csv::emit_parse_line_body(
            std::slice::from_mut(chunk),
            0,
            slot,
            slot + 1,
            slot + 2,
            line,
        ),
        "__stdlib_csv_parse_document" => crate::primitives::csv::emit_parse_document_body(
            std::slice::from_mut(chunk),
            0,
            slot,
            slot + 1,
            slot + 2,
            line,
        ),
        "__stdlib_csv_format_minimal" | "__stdlib_csv_format_whitespace" => {
            crate::primitives::csv::emit_format_row_body(
                std::slice::from_mut(chunk),
                0,
                slot,
                slot + 1,
                slot + 2,
                crate::primitives::csv::FormatOptions {
                    quote_whitespace: export == "__stdlib_csv_format_whitespace",
                },
                line,
            )
        }
        _ => unreachable!("private inline body was not registered"),
    }
    assert_eq!(
        metadata_sizes,
        (
            chunk.call_indirect_sigs.len(),
            chunk.call_indirect_canon.len(),
            chunk.stack_switch_handlers.len(),
            chunk.direct_defined_calls.len()
        ),
        "private inline body added offset metadata requiring relocation"
    );
    assert_eq!(
        original_private_refs,
        chunk
            .local_names
            .iter()
            .filter(|local| private_ref_marker(&local.name).is_some())
            .map(|local| (local.slot, local.name.clone()))
            .collect::<Vec<_>>(),
        "private inline body introduced pending private calls"
    );
    assert!(
        original_i32_blocks.is_subset(&chunk.block_i32_results),
        "private inline body modified existing I32 block metadata"
    );
    let appended_i32_blocks: Vec<_> = chunk
        .block_i32_results
        .difference(&original_i32_blocks)
        .copied()
        .collect();
    if export == "__stdlib_csv_parse_document" {
        assert_eq!(
            appended_i32_blocks.len(),
            1,
            "CSV Document must add exactly its trailing I32 IF"
        );
    } else {
        assert!(
            appended_i32_blocks.is_empty(),
            "private inline body added unexpected I32 block metadata"
        );
    }
    let appended_boundaries = instruction_boundaries(chunk);
    for &key in &appended_i32_blocks {
        assert!(
            key >= original_len && key < chunk.code.len() && appended_boundaries.contains(&key),
            "appended I32 block is not an instruction boundary"
        );
        assert_eq!(
            Op::decode(chunk.read_u16(key), chunk.read_u16(key + 2)),
            Some(Op::IF)
        );
        assert_eq!(
            &chunk.code[key + 4..key + 6],
            &[0, 1],
            "appended CSV predicate must be a zero-param single-result IF"
        );
    }
    let replacement = chunk.code.split_off(original_len);
    let replacement_lines = chunk.lines.split_off_table(original_len);
    assert_eq!(replacement.len(), replacement_lines.len());
    let replacement_len = replacement.len();
    chunk.block_i32_results = original_i32_blocks
        .into_iter()
        .map(|key| remap_offset(key, offset, end, replacement_len))
        .collect();
    for key in appended_i32_blocks {
        let relocated = offset
            .checked_add(key - original_len)
            .expect("appended I32 block offset overflow");
        assert!(
            chunk.block_i32_results.insert(relocated),
            "duplicate relocated I32 block metadata"
        );
    }
    chunk.direct_defined_calls = std::mem::take(&mut chunk.direct_defined_calls)
        .into_iter().map(|(key,value)|(remap_offset(key,offset,end,replacement_len),value)).collect();
    chunk.call_indirect_sigs = std::mem::take(&mut chunk.call_indirect_sigs)
        .into_iter()
        .map(|(key, value)| (remap_offset(key, offset, end, replacement_len), value))
        .collect();
    chunk.call_indirect_canon = std::mem::take(&mut chunk.call_indirect_canon)
        .into_iter()
        .map(|(key, value)| (remap_offset(key, offset, end, replacement_len), value))
        .collect();
    chunk.stack_switch_handlers = std::mem::take(&mut chunk.stack_switch_handlers)
        .into_iter()
        .map(|(key, mut handlers)| {
            for handler in &mut handlers {
                // The sole source compiler producer (generators::emit_next)
                // emits a forward absolute byte offset. Standard WASM readers
                // emit label depths, which cannot exceed the opcode offset.
                // Preserve those depths; relocate only proven forward targets.
                if handler.kind == 0 && handler.label_index as usize > key {
                    let target = handler.label_index as usize;
                    assert!(
                        boundaries.contains(&target),
                        "forward generator handler is not an instruction boundary"
                    );
                    handler.label_index =
                        u32::try_from(remap_offset(target, offset, end, replacement_len))
                            .expect("generator handler offset exceeds u32");
                }
            }
            (remap_offset(key, offset, end, replacement_len), handlers)
        })
        .collect();
    if let Some(user) = chunk.user_code_offset {
        // The source entry boundary may coincide with the first replaced op.
        chunk.user_code_offset = Some(if user as usize == offset {
            user
        } else {
            u32::try_from(remap_offset(user as usize, offset, end, replacement_len))
                .expect("user code offset exceeds u32")
        });
    }
    generator_native_markers::relocate(&mut chunk.local_names, |p| {
        remap_offset(p, offset, end, replacement_len)
    });
    chunk.code.splice(offset..end, replacement);
    chunk.lines.replace_range_table(offset..end, replacement_lines);
    let final_boundaries = instruction_boundaries(chunk);
    assert!(
        chunk
            .block_i32_results
            .iter()
            .all(|key| final_boundaries.contains(key)),
        "relocated I32 block is not an instruction boundary"
    );
}

fn resolve_private_helper_refs(chunks: &mut [Chunk], inline_exports: &BTreeSet<&'static str>) {
    use vybe_runtime::chunk::LocalOrigin;
    let indices: std::collections::BTreeMap<_, _> = PRIVATE_HELPERS
        .iter()
        .filter_map(|&(export, _)| {
            private_helper_index(chunks, export).map(|index| (export, index))
        })
        .collect();
    for chunk in chunks {
        let mut refs = std::collections::BTreeMap::new();
        for local in &chunk.local_names {
            if local.origin != LocalOrigin::Compiler {
                continue;
            }
            if let Some((export, offset)) = private_ref_marker(&local.name) {
                if let Some(previous) = refs.insert(offset, (export, local.slot)) {
                    assert_eq!(
                        previous,
                        (export, local.slot),
                        "conflicting private reference records"
                    );
                }
            }
        }
        // Highest offsets first: inlining cannot invalidate earlier records.
        for (offset, (export, slot)) in refs.into_iter().rev() {
            assert!(
                offset
                    .checked_add(7)
                    .is_some_and(|end| end <= chunk.code.len()),
                "private helper reference offset is out of bounds"
            );
            if private_range_direct(export) && Op::decode(chunk.read_u16(offset), chunk.read_u16(offset + 2)) == Some(Op::CALL) {
                let edge = chunk.direct_defined_calls.get_mut(&offset).expect("private direct-call edge missing");
                assert_eq!(edge.target, u32::MAX, "private direct target changed before finalization");
                assert!(!edge.source);
                assert_eq!(edge.params, "externref,externref");
                assert_eq!(edge.results, "externref");
                edge.target = u32::try_from(*indices.get(export).expect("private direct helper was not linked")).expect("Wasm function index exceeds u32");
                continue;
            }
            assert_eq!(
                Op::decode(chunk.read_u16(offset), chunk.read_u16(offset + 2)),
                Some(Op::REF_FUNC),
                "private helper reference offset does not point to REF_FUNC"
            );
            assert_eq!(
                chunk.read_u16(offset + 4),
                UNRESOLVED_PRIVATE_REF,
                "private helper reference was modified before finalization"
            );
            assert_eq!(
                chunk.code[offset + 6],
                0,
                "private helper references cannot capture upvalues"
            );
            if inline_exports.contains(export)
                || indices
                    .get(export)
                    .is_some_and(|&index| index > u16::MAX as usize)
            {
                inline_private_call(chunk, export, offset, slot);
            } else {
                let index =
                    u16::try_from(*indices.get(export).expect("private helper was not linked"))
                        .expect("private helper function index exceeds the u16 REF_FUNC encoding");
                let [hi, lo] = index.to_be_bytes();
                chunk.code[offset + 4] = hi;
                chunk.code[offset + 5] = lo;
            }
        }
        chunk.local_names.retain(|local| {
            local.origin != LocalOrigin::Compiler || private_ref_marker(&local.name).is_none()
        });
    }
}

/// Mapping from helper chunk name to the global name used at call sites.
const MAPPINGS: &[(&str, &str)] = &[
    ("__stdlib_reverse_copy", "__vybe_reverse_copy"),
    (
        "__stdlib_reverse_copy_own_fields",
        "__vybe_reverse_copy_own_fields",
    ),
    (
        "__stdlib_reverse_copy_prototype_members",
        "__vybe_reverse_copy_prototype_members",
    ),
    (
        "__stdlib_reverse_copy_own_fields_prototype_members",
        "__vybe_reverse_copy_own_fields_prototype_members",
    ),
    (
        "__stdlib_bytes_to_binary_string_all",
        "__vybe_bytes_to_binary_string_all",
    ),
    (
        "__stdlib_bytes_to_binary_string",
        "__vybe_bytes_to_binary_string",
    ),
    (
        "__stdlib_binary_string_to_bytes",
        "__vybe_binary_string_to_bytes",
    ),
    ("__stdlib_autoderef", "__vybe_autoderef"),
    ("__stdlib_dyneq", "__vybe_dyneq"),
    ("__stdlib_dynlt", "__vybe_dynlt"),
    ("__stdlib_dyngt", "__vybe_dyngt"),
    ("__stdlib_dynle", "__vybe_dynle"),
    ("__stdlib_dynge", "__vybe_dynge"),
    (
        "__stdlib_encoding_base64_decode",
        "__vybe_encoding_base64_decode",
    ),
    ("__stdlib_binary_string_to_uint8_array", "__vybe_binary_string_to_uint8_array"),
    ("__stdlib_base64_decode_bytes", "__vybe_base64_decode_bytes"),
    ("__stdlib_encoding_bin2hex", "__vybe_encoding_bin2hex"),
    ("__stdlib_encoding_hex2bin", "__vybe_encoding_hex2bin"),
    ("__stdlib_encoding_hex2bytes", "__vybe_encoding_hex2bytes"),
    ("__stdlib_encoding_str_rot13", "__vybe_encoding_str_rot13"),
    (
        "__stdlib_encoding_quoted_printable_encode",
        "__vybe_encoding_quoted_printable_encode",
    ),
    (
        "__stdlib_encoding_quoted_printable_decode",
        "__vybe_encoding_quoted_printable_decode",
    ),
    (
        "__stdlib_encoding_convert_uuencode",
        "__vybe_encoding_convert_uuencode",
    ),
    (
        "__stdlib_encoding_convert_uudecode",
        "__vybe_encoding_convert_uudecode",
    ),
    (
        "__stdlib_encoding_base64_decode_lenient",
        "__vybe_encoding_base64_decode_lenient",
    ),
    ("__stdlib_sort_by_key", "__vybe_sort_by_key"),
    (
        "__stdlib_sort_by_key_receiver",
        "__vybe_sort_by_key_receiver",
    ),
    ("__stdlib_heapify", "__vybe_heapify"),
    ("__stdlib_nsmallest", "__vybe_nsmallest"),
    ("__stdlib_nlargest", "__vybe_nlargest"),
    ("__stdlib_nsmallest_key", "__vybe_nsmallest_key"),
    ("__stdlib_nlargest_key", "__vybe_nlargest_key"),
    (
        "__stdlib_nsmallest_key_receiver",
        "__vybe_nsmallest_key_receiver",
    ),
    (
        "__stdlib_nlargest_key_receiver",
        "__vybe_nlargest_key_receiver",
    ),
    ("__stdlib_sorted", "__vybe_sorted"),
    ("__stdlib_chan_send", "__vybe_chan_send"),
    ("__stdlib_chan_recv", "__vybe_chan_recv"),
    ("__stdlib_chan_recv_ok", "__vybe_chan_recv_ok"),
    ("__stdlib_chan_len", "__vybe_chan_len"),
    ("__stdlib_chan_cap", "__vybe_chan_cap"),
    ("__stdlib_chan_close", "__vybe_chan_close"),
    ("__stdlib_chan_ready_recv", "__vybe_chan_ready_recv"),
    ("__stdlib_chan_ready_send", "__vybe_chan_ready_send"),
    ("__stdlib_chan_wait_slice", "__vybe_chan_wait_slice"),
    ("__stdlib_chan_try_send", "__vybe_chan_try_send"),
    ("__stdlib_chan_try_recv", "__vybe_chan_try_recv"),
    ("__stdlib_chan_try_peek", "__vybe_chan_try_peek"),
    ("__stdlib_chan_drained", "__vybe_chan_drained"),
    ("__stdlib_chan_closed", "__vybe_chan_closed"),
    ("__stdlib_chan_recv_or_throw", "__vybe_chan_recv_or_throw"),
    ("__stdlib_chan_wait_readable", "__vybe_chan_wait_readable"),
    ("__stdlib_futex_alloc16", "__vybe_futex_alloc16"),
    ("__stdlib_task_new", "__vybe_task_new"),
    ("__stdlib_task_wait", "__vybe_task_wait"),
    ("__stdlib_sort_in_place", "__vybe_sort_in_place"),
    (
        "__stdlib_sort_with_comparator",
        "__vybe_sort_with_comparator",
    ),
    // C qsort/bsearch moved to platforms/libc/stdlib_runtime.rs (libc surface);
    // no longer part of the cross-language helper bundle.
    ("__stdlib_enumerate", "__vybe_enumerate"),
    ("__stdlib_sum", "__vybe_sum"),
    ("__stdlib_min", "__vybe_min"),
    ("__stdlib_max", "__vybe_max"),
    ("__stdlib_pyany", "__vybe_pyany"),
    ("__stdlib_pyall", "__vybe_pyall"),
    ("__stdlib_compact", "__vybe_compact"),
    ("__stdlib_isempty", "__vybe_isempty"),
    ("__stdlib_pymap", "__vybe_pymap"),
    ("__stdlib_pyfilter", "__vybe_pyfilter"),
    ("__stdlib_pyiter", "__vybe_pyiter"),
    ("__stdlib_pynext", "__vybe_pynext"),
    ("__stdlib_array_copy", "__vybe_array_copy"),
    // Math transcendentals removed — provided natively by `ecma:math:*`.
    // `__stdlib_tostring` removed — str()/toString route to `ecma:string.String`.
    (
        "__stdlib_string_is_null_or_empty",
        "__vybe_string_is_null_or_empty",
    ),
    (
        "__stdlib_string_is_null_or_whitespace",
        "__vybe_string_is_null_or_whitespace",
    ),
    ("__stdlib_str_insert", "__vybe_str_insert"),
    ("__stdlib_str_remove_start", "__vybe_str_remove_start"),
    ("__stdlib_str_remove_range", "__vybe_str_remove_range"),
    ("__stdlib_pascal_set_include", "__vybe_pascal_set_include"),
    ("__stdlib_pascal_set_exclude", "__vybe_pascal_set_exclude"),
    ("__stdlib_pascal_set_union", "__vybe_pascal_set_union"),
    (
        "__stdlib_pascal_set_intersection",
        "__vybe_pascal_set_intersection",
    ),
    (
        "__stdlib_pascal_set_difference",
        "__vybe_pascal_set_difference",
    ),
    ("__stdlib_pascal_set_contains", "__vybe_pascal_set_contains"),
    ("__stdlib_pascal_str_insert", "__vybe_pascal_str_insert"),
    (
        "__stdlib_pascal_str_remove_range",
        "__vybe_pascal_str_remove_range",
    ),
    ("__stdlib_isnumeric", "__vybe_isnumeric"),
    ("__stdlib_val", "__vybe_val"),
    ("__stdlib_iif", "__vybe_iif"),
    ("__stdlib_rgb", "__vybe_rgb"),
    ("__stdlib_qbcolor", "__vybe_qbcolor"),
    ("__stdlib_isdate", "__vybe_isdate"),
    ("__stdlib_vartype", "__vybe_vartype"),
    ("__stdlib_newline", "__vybe_newline"),
    (
        "__stdlib_dict_values_from_entries",
        "__vybe_dict_values_from_entries",
    ),
    ("__stdlib_setdefault", "__vybe_setdefault"),
    ("__stdlib_to_bytes", "__vybe_to_bytes"),
    ("__stdlib_id", "__vybe_id"),
    ("__stdlib_hash", "__vybe_hash"),
    (
        "__stdlib_dotnet_numeric_format",
        "__vybe_dotnet_numeric_format",
    ),
    ("__stdlib_format_map", "__vybe_format_map"),
    ("__stdlib_pyhex", "__vybe_pyhex"),
    ("__stdlib_pyoct", "__vybe_pyoct"),
    ("__stdlib_pybin", "__vybe_pybin"),
    ("__stdlib_isinf", "__vybe_isinf"),
    // `__stdlib_splice` removed — no emit site references `__vybe_splice`.
    ("__stdlib_js_get_method", "__vybe_js_get_method"),
    ("__stdlib_redim", "__vybe_redim"),
    ("__stdlib_slicestep", "__vybe_slicestep"),
    ("__stdlib_dynmul", "__vybe_dynmul"),
    ("__stdlib_fmod", "__vybe_fmod"),
    ("__stdlib_array_insert", "__vybe_array_insert"),
    ("__stdlib_array_remove_at", "__vybe_array_remove_at"),
    ("__stdlib_array_remove_value", "__vybe_array_remove_value"),
    ("__stdlib_array_insert_range", "__vybe_array_insert_range"),
    ("__stdlib_array_set_range", "__vybe_array_set_range"),
    ("__stdlib_array_binary_search", "__vybe_array_binary_search"),
    ("__stdlib_array_reverse_range", "__vybe_array_reverse_range"),
    ("__stdlib_generator_next", "__vybe_generator_next"),
    (
        "__stdlib_async_generator_next",
        "__vybe_async_generator_next",
    ),
    ("__stdlib_generator_self", "__vybe_generator_self"),
    ("__stdlib_iter_drain", "__vybe_iter_drain"),
    // PHP runtime helpers — all inline opcode emitters under
    // `emitter/php/<cat>_adapter.rs`. Reached through `common:php.*`
    // dispatch arms; no `__vybe_*` global indirection.
    // PHP filesystem helpers — inline opcode emitters under
    // `emitter/php/filesystem_adapter.rs`. Reached via `common:php.*`
    // dispatch arms; no `__vybe_*` global indirection.
    // Regex pattern-first adapters (PHP preg_*, Python re.*, VB Regex.*)
    (
        "__stdlib_regex_replace_pat_first",
        "__ecma_regexp_replace_pat_first",
    ),
    (
        "__stdlib_regex_split_pat_first",
        "__ecma_regexp_split_pat_first",
    ),
    (
        "__stdlib_regex_match_all_pat_first",
        "__ecma_regexp_match_all_pat_first",
    ),
];

fn mappings() -> impl Iterator<Item = &'static (&'static str, &'static str)> {
    MAPPINGS
        .iter()
        .chain(crate::primitives::array_transforms::SHARED_HELPERS.iter())
}

// These recipes are pure fixed-ABI bodies, independent of source scope and
// language profile. Identity is process-local authority minted by the Rust
// producer, not a guest name or a serialized/debug-metadata spelling.
fn fixed_runtime_providers() -> &'static [(&'static str, vybe_runtime::shared_provider::RuntimeHelperReceipt)] {
    static PROVIDERS: OnceLock<Vec<(&'static str, vybe_runtime::shared_provider::RuntimeHelperReceipt)>> = OnceLock::new();
    PROVIDERS.get_or_init(|| [
        ("__stdlib_dyneq", "__vybe_dyneq", 2, "vybe.compiler.fixed-comparison.v1"),
        ("__stdlib_dynlt", "__vybe_dynlt", 2, "vybe.compiler.fixed-comparison.v1"),
        ("__stdlib_dyngt", "__vybe_dyngt", 2, "vybe.compiler.fixed-comparison.v1"),
        ("__stdlib_dynle", "__vybe_dynle", 2, "vybe.compiler.fixed-comparison.v1"),
        ("__stdlib_dynge", "__vybe_dynge", 2, "vybe.compiler.fixed-comparison.v1"),
        // The standalone reference reader uses only its argument, fixed
        // property keys and runtime imports. It owns no class/type indices,
        // source bindings, directive-dependent policy or captured values.
        // Reuse the algorithm, never a previously read reference value.
        ("__stdlib_autoderef", "__vybe_autoderef", 1, "vybe.compiler.fixed-autoderef.v1"),
    ].into_iter().map(|(export, global, arity, version)| (export,
        vybe_runtime::shared_provider::RuntimeHelperReceipt::new(global,
            version, arity, 1))).collect())
}

pub(crate) fn snapshot_runtime_helpers(vm: &vybe_runtime::VM)
    -> Vec<vybe_runtime::shared_provider::InstalledRuntimeProvider>
{
    fixed_runtime_providers().iter()
        .filter_map(|(_, receipt)| vm.snapshot_runtime_helper(receipt))
        .filter(|provider| runtime_provider_diagnostics_clean(vm, provider)).collect()
}

pub(crate) fn runtime_provider_diagnostics_clean(vm: &vybe_runtime::VM,
    provider: &vybe_runtime::shared_provider::InstalledRuntimeProvider) -> bool {
    let Some(chunk) = vm.chunks.get(provider.chunk_index()) else { return false; };
    let mut ip = 0;
    while ip < chunk.code.len() {
        let op = Op::from_code_at(&chunk.code, ip);
        if op == Op::CALL && !chunk.direct_defined_calls.contains_key(&ip) {
            let import = u16::from_be_bytes(chunk.code[ip+4..ip+6].try_into().unwrap()) as usize;
            // Linking moves a function's imports onto its owning module.
            // Resolve the descriptor through the attested live provider;
            // an empty function-local table is not a missing host import.
            let Some(import) = vm.runtime_helper_import(provider, import) else { return false; };
            if vybe_runtime::host_arity_mismatch(&import.module, &import.name, chunk.code[ip+6]).is_some() {
                // Keep the original builder/diagnostic path when the live
                // registry would report an arity mismatch during construction.
                return false;
            }
        }
        ip += 4 + op.operand_format().size_in(&chunk.code, ip+4);
    }
    true
}

pub(crate) fn installed_runtime_helper_exports(
    providers: &[vybe_runtime::shared_provider::InstalledRuntimeProvider],
    chunks: &[Chunk],
) -> Vec<&'static str> {
    fixed_runtime_providers().iter().filter_map(|(export, receipt)| {
        (providers.iter().any(|provider| provider.has_receipt(receipt))
            && !chunks.iter().any(|chunk| chunk.global_inits.iter()
                .any(|init| init.name == receipt.global_name())))
            .then_some(*export)
    }).collect()
}

/// Emit the runtime-helper preamble at the START of a script chunk.
/// This must be called BEFORE any user code is emitted.
///
/// The preamble emits, for each helper function:
///   `if globals[name] is null { globals[name] = ref_func(stdlib_chunk) }`
///
/// This is the polyfill pattern: if the host (Vybe VM) has already populated
/// `__vybe_*` globals with optimized native fns BEFORE running the script,
/// the preamble leaves those alone. On non-Vybe runtimes the globals start
/// null and the preamble installs the bundled helper bytecode chunks.
pub fn emit_runtime_helper_preamble(script: &mut Chunk, stdlib_base: usize) {
    for (i, &(_chunk_name, global_name)) in mappings().enumerate() {
        let ci = stdlib_base + i;

        // Check if global is already set: global_get + ref_is_null
        crate::primitives::globals::emit_read(script, global_name, 0);
        script.emit_op(Op::REF_IS_NULL, 0);
        // `[is_null]` in, nothing out: the install path pushes the funcref
        // and `GLOBAL_SET` consumes it. One param, no result.
        let install_block = script.emit_block_params(0, 1, 0);
        script.emit_op(Op::I32_EQZ, 0);
        script.emit_br_if(0, 0);

        // Install: ref_func + global_set + drop
        script.emit_op_u16(Op::REF_FUNC, ci as u16, 0);
        script.emit(0, 0); // 0 upvalues
        crate::primitives::globals::emit_write(script, global_name, 0);

        script.emit_end(0);
        script.patch_block(install_block);
    }
}

/// Emit a call to a runtime-helper/vybe function.
/// IMPORTANT: push the function ref FIRST (via global_get), then push args, then call_ref.
/// The call convention is: [func_ref, arg0, arg1, ...] on stack.
///
/// Usage in compiler:
///   emit_call_push_func(chunk, "__vybe_sorted", 0);  // push func ref
///   compile_expr(arg0);                               // push array
///   emit_call_invoke(chunk, 1, 0);                    // call_ref 1
pub fn emit_call_push_func(chunk: &mut Chunk, global_name: &str, line: u32) {
    crate::primitives::globals::emit_read(chunk, global_name, line);
}

/// Emit call_ref after func + args are on stack.
pub fn emit_call_invoke(chunk: &mut Chunk, argc: u8, line: u32) {
    crate::primitives::callable::emit_direct_invoke_chunk(chunk, argc, line);
}

/// Append referenced runtime helper chunks and register them as global_inits.
/// Call this at the END of compilation, after all user chunks are finalized.
/// The script chunk (chunks[0]) gets global_inits with RefFunc entries.
pub fn finalize_with_runtime_helpers(chunks: &mut Vec<Chunk>) {
    finalize_with_runtime_helpers_excluding(chunks, &[]);
}

/// Same as [`finalize_with_runtime_helpers`], but skips selected helper exports.
/// Used by profiles with platform-owned replacements for a generic helper.
pub fn finalize_with_runtime_helpers_excluding(chunks: &mut Vec<Chunk>, excluded_exports: &[&str]) {
    use vybe_runtime::chunk::{ConstExpr, GlobalInit};

    let mut requested = referenced_helper_exports(chunks);
    add_helper_dependencies(&mut requested);
    for excluded in excluded_exports {
        if private_helper_arity(excluded).is_none() {
            requested.remove(*excluded);
        }
    }
    if requested.is_empty() {
        return;
    }

    // Only treat a body as installed when its global declaration actually
    // points to the expected helper chunk. Finalization can run again after
    // linking modules; that must not duplicate previously bundled algorithms.
    let mut installed: BTreeSet<_> = chunks[0]
        .global_inits
        .iter()
        .filter_map(|init| {
            let export = helper_export_for_global(&init.name)?;
            let ConstExpr::RefFunc(index) = &init.init else {
                return None;
            };
            (chunks.get(*index)?.name == export).then_some(export)
        })
        .collect();
    installed.extend(
        PRIVATE_HELPERS
            .iter()
            .filter_map(|&(export, _)| private_helper_index(chunks, export).map(|_| export)),
    );
    let helper_base = chunks.len();
    // Build helper chunks with their imports registered on `chunks[0]`
    // (the module-level imports section — single per WASM module).
    // Same dependency surface as user code → helpers become true
    // cross-runtime polyfill: on Vybe the chunks are swapped for
    // native handlers; on v8 their `ecma:array.*` imports resolve
    // to native `Array.prototype.*`; on wasmtime the Phase-C polyfill
    // supplies the imports.
    let mut helper_chunks = Vec::new();
    let mut helper_exports = Vec::new();
    let mut inline_exports = BTreeSet::new();
    // A single similar_text site gains no body sharing. Inline its optimized
    // primitive to remove the extra function and call frame, provided the
    // caller can accommodate its six private locals. Repeated sites retain
    // one shared body; the existing wide-index fallback uses the same path.
    if !installed.contains("__stdlib_similar_text") {
        let sites: BTreeSet<_> = chunks.iter().enumerate().flat_map(|(index, chunk)| {
            chunk.local_names.iter().filter_map(move |local| {
                (local.origin == vybe_runtime::chunk::LocalOrigin::Compiler)
                    .then(|| private_ref_marker(&local.name)).flatten()
                    .filter(|(export, _)| *export == "__stdlib_similar_text")
                    .map(|(_, offset)| (index, offset))
            })
        }).collect();
        if sites.len() == 1 {
            let (index, _) = *sites.iter().next().unwrap();
            if chunks[index].local_count <= u16::MAX - 6 {
                inline_exports.insert("__stdlib_similar_text");
            }
        }
    }
    requested.retain(|name| !inline_exports.contains(name));
    loop {
        let candidates: Vec<_> = ordered_helper_exports(&requested)
            .into_iter()
            .filter(|name| !installed.contains(name) && !helper_exports.contains(name))
            .collect();
        let mut missing = Vec::new();
        for name in candidates {
            let index = helper_base + helper_chunks.len() + missing.len();
            if private_helper_arity(name).is_some() && !private_range_direct(name) && index > u16::MAX as usize {
                inline_exports.insert(name);
            } else {
                missing.push(name);
            }
        }
        requested.retain(|name| !inline_exports.contains(name));
        if missing.is_empty() {
            break;
        }
        let mut built = Vec::new();
        for &name in &missing {
            if private_helper_arity(name).is_some() {
                built.push(build_private_primitive(name).expect("registered primitive body"));
            } else {
                // Existing public helper definitions remain available until
                // their callers migrate to the relevant adapter/primitive.
                built.extend(crate::primitives::polyfills::build_runtime_helpers_for_exports(
                    &mut chunks[0], &[name],
                ).chunks);
            }
        }
        requested.extend(referenced_helper_exports(&built));
        // Builders can register references on the module-level import carrier.
        requested.extend(referenced_helper_exports(&chunks[..1]));
        add_helper_dependencies(&mut requested);
        requested.retain(|name| {
            !inline_exports.contains(name)
                && (private_helper_arity(name).is_some() || !excluded_exports.contains(name))
        });
        helper_chunks.extend(built);
        helper_exports.extend(missing);
    }

    for (i, &chunk_name) in helper_exports.iter().enumerate() {
        if let Some(global_name) = helper_global_for_export(chunk_name) {
            chunks[0].global_inits.push(GlobalInit {
                name: global_name.to_string(),
                init: ConstExpr::RefFunc(helper_base + i),
            });
        }
    }
    for chunk in &mut helper_chunks {
        if let Some((_, receipt)) = fixed_runtime_providers().iter()
            .find(|(export, _)| *export == chunk.name)
        {
            chunk.runtime_helper_receipt = Some(receipt.clone());
        }
    }
    chunks.extend(helper_chunks);
    resolve_private_helper_refs(chunks, &inline_exports);
}

/// Convenience: emit func_ref push + args already on stack + call_ref.
/// Args must already be on stack BEFORE calling this.
/// This function inserts the func ref below the args using a temp local approach.
/// Actually — we can't easily insert below stack items without a local.
/// Instead, callers should use emit_call_push_func + args + emit_call_invoke.
///
/// For backward compatibility, this emits global_get + call_ref, but callers
/// MUST push args AFTER calling this (which means this only works for 0-arg calls).
/// For multi-arg calls, use the push_func/invoke pair.
pub fn emit_call(chunk: &mut Chunk, global_name: &str, argc: u8, line: u32) {
    crate::primitives::globals::emit_read(chunk, global_name, line);
    crate::primitives::callable::emit_direct_invoke_chunk(chunk, argc, line);
}

/// Outline a body while preserving an already evaluated argument stack.
/// The helper's explicit arity is independent of its user callback receiver ABI.
pub(crate) fn emit_stacked_helper_call(c: &mut Chunk, global: &str, argc: u8, line: u32) {
    let base = c.alloc_scratch(argc as u16);
    for offset in (0..argc).rev() {
        c.emit_op_u16(Op::LOCAL_SET, base + offset as u16, line);
    }
    crate::primitives::globals::emit_read(c, global, line);
    for offset in 0..argc {
        c.emit_op_u16(Op::LOCAL_GET, base + offset as u16, line);
    }
    crate::primitives::callable::emit_direct_invoke_chunk(c, argc, line);
}

fn referenced_helper_exports(chunks: &[Chunk]) -> BTreeSet<&'static str> {
    let mut exports = BTreeSet::new();
    for chunk in chunks {
        for local in &chunk.local_names {
            if local.origin == vybe_runtime::chunk::LocalOrigin::Compiler {
                if let Some((export, _)) = private_ref_marker(&local.name) {
                    exports.insert(export);
                }
            }
        }
        for constant in &chunk.constants {
            if let Value::String(name) = constant {
                if let Some(export) = helper_export_for_global(name.as_ref()) {
                    exports.insert(export);
                }
            }
        }
    }
    // Any program with a generator needs the continuation protocol helpers.
    // `attach_continuation_protocols` (VM) wires `__vybe_generator_next` /
    // `__vybe_generator_self` onto every continuation as its real `next` /
    // `[Symbol.iterator]` methods — `__stdlib_generator_next` is a pure
    // bytecode driver (`emit_next` → spec `resume`). Generators drive via
    // inline `emit_next`, so nothing *references* these globals by name; force
    // them in so a generator is a first-class iterator (`it.next` is readable
    // and callable) for generic for-of / `Iterator.from` / spread paths.
    if chunks.iter().any(|chunk| chunk.is_generator) {
        exports.insert("__stdlib_generator_next");
        exports.insert("__stdlib_generator_self");
        // §27.6.1.2: async generators' `next()` returns a PROMISE-wrapped
        // IteratorResult — attach picks the async driver for them.
        exports.insert("__stdlib_async_generator_next");
    }
    exports
}

fn add_helper_dependencies(exports: &mut BTreeSet<&'static str>) {
    loop {
        let before = exports.len();
        for export in exports.clone() {
            for dep in helper_export_dependencies(export) {
                exports.insert(dep);
            }
        }
        if exports.len() == before {
            break;
        }
    }
}

fn helper_export_dependencies(export: &str) -> &'static [&'static str] {
    match export {
        "__stdlib_pynext" => &["__stdlib_iter_drain"],
        _ => &[],
    }
}

fn ordered_helper_exports(exports: &BTreeSet<&'static str>) -> Vec<&'static str> {
    mappings()
        .filter_map(|&(chunk_name, _global_name)| {
            exports.contains(chunk_name).then_some(chunk_name)
        })
        .chain(
            PRIVATE_HELPERS
                .iter()
                .filter_map(|&(export, _)| exports.contains(export).then_some(export)),
        )
        .collect()
}

fn helper_export_for_global(global_name: &str) -> Option<&'static str> {
    if let Some(&(export, global)) = MAPPINGS.first() {
        if global_name == global {
            return Some(export);
        }
    }
    helper_mapping_indexes().by_global.get(global_name).copied()
}

fn helper_global_for_export(export: &str) -> Option<&'static str> {
    if let Some(&(first_export, global)) = MAPPINGS.first() {
        if export == first_export {
            return Some(global);
        }
    }
    helper_mapping_indexes().by_export.get(export).copied()
}

struct HelperMappingIndexes {
    by_global: HashMap<&'static str, &'static str>,
    by_export: HashMap<&'static str, &'static str>,
}

fn helper_mapping_indexes() -> &'static HelperMappingIndexes {
    static INDEXES: OnceLock<HelperMappingIndexes> = OnceLock::new();
    INDEXES.get_or_init(|| {
        let mut by_global = HashMap::new();
        let mut by_export = HashMap::new();
        for &(export, global) in mappings() {
            // Keep the ordered catalog's first match for duplicate names.
            by_global.entry(global).or_insert(export);
            by_export.entry(export).or_insert(global);
        }
        HelperMappingIndexes {
            by_global,
            by_export,
        }
    })
}

#[cfg(test)]
mod helper_mapping_lookup_tests {
    use super::*;

    #[test]
    fn linked_fixed_provider_uses_module_import_owner_and_keeps_diagnostics() {
        let mut script = Chunk::new("<script>");
        emit_call_push_func(&mut script, "__vybe_autoderef", 0);
        script.emit_f64_const(7.0, 0);
        emit_call_invoke(&mut script, 1, 0);
        script.emit_op(Op::RETURN, 0);
        let mut chunks = vec![script];
        finalize_with_runtime_helpers(&mut chunks);
        let index = chunks.iter().position(|c| c.name == "__stdlib_autoderef").unwrap();

        // Model normal linked-module import ownership without reconstructing
        // descriptor identities from names or publishing private VM tables.
        let base = chunks[0].imports.len();
        let imports = std::mem::take(&mut chunks[index].imports);
        assert!(!imports.is_empty());
        chunks[0].imports.extend(imports);
        let helper = &mut chunks[index];
        let mut ip = 0;
        while ip < helper.code.len() {
            let op = Op::from_code_at(&helper.code, ip);
            if op == Op::CALL && !helper.direct_defined_calls.contains_key(&ip) {
                let old = u16::from_be_bytes(helper.code[ip+4..ip+6].try_into().unwrap());
                let relocated = u16::try_from(base + old as usize).unwrap();
                helper.code[ip+4..ip+6].copy_from_slice(&relocated.to_be_bytes());
            }
            ip += 4 + op.operand_format().size_in(&helper.code, ip+4);
        }
        crate::primitives::globals::normalize_global_table(&mut chunks);
        let mut vm = vybe_runtime::VM::new();
        crate::primitives::platforms::register_platforms_all(&mut vm);
        let result = vm.run(chunks).unwrap();
        assert!(matches!(result, Value::F64(value) if value == 7.0));
        assert!(vm.chunks[index].imports.is_empty());
        let receipt = vm.chunks[index].runtime_helper_receipt.clone().unwrap();
        let installed = vm.snapshot_runtime_helper(&receipt).unwrap();
        assert!(runtime_provider_diagnostics_clean(&vm, &installed));

        let mut certificate = crate::compile_dependencies::Certificate::default();
        certificate.install_runtime_helpers(snapshot_runtime_helpers(&vm));
        let mut consumer = Chunk::new("<script>");
        emit_call_push_func(&mut consumer, "__vybe_autoderef", 0);
        consumer.emit_op(Op::RETURN, 0);
        let mut included = vec![consumer];
        let excluded = certificate.runtime_helper_exports(&included);
        assert!(excluded.contains(&"__stdlib_autoderef"));
        finalize_with_runtime_helpers_excluding(&mut included, &excluded);
        assert_eq!(included.len(), 1);

        // Standard imports can legitimately have no process-global arity
        // declaration. Use a unique declared contract for the negative case,
        // rather than treating an unknown signature as a mismatch or mutating
        // another test's platform registrations. The bad call is never run.
        vybe_runtime::declare_host_signature("test:linked-provider-arity", "unary",
            vybe_runtime::FuncSig {
                name: "unary".into(),
                params: vybe_runtime::Param::unnamed_list(vec![vybe_runtime::ValType::String]),
                results: vec![],
            }, None);
        assert_eq!(vybe_runtime::host_arity_mismatch("test:linked-provider-arity", "unary", 2), Some(1));
        let bad_import = vm.chunks[0].add_import("test:linked-provider-arity", "unary");
        let mut ip = 0;
        let mut mismatched = false;
        while ip < vm.chunks[index].code.len() {
            let helper = &vm.chunks[index];
            let op = Op::from_code_at(&helper.code, ip);
            if op == Op::CALL && !helper.direct_defined_calls.contains_key(&ip) {
                vm.chunks[index].code[ip+4..ip+6].copy_from_slice(&bad_import.to_be_bytes());
                vm.chunks[index].code[ip+6] = 2;
                mismatched = true;
                break;
            }
            ip += 4 + op.operand_format().size_in(&helper.code, ip+4);
        }
        assert!(mismatched, "linked provider must exercise the declared diagnostic veto");
        assert!(!runtime_provider_diagnostics_clean(&vm, &installed));
        assert!(!certificate.runtime_helpers_valid(&vm));
    }

    #[test]
    fn fixed_provider_reuse_requires_exact_live_request_owned_authority() {
        use vybe_runtime::value::{Function, Object, ObjectKind};
        for &(export, ref receipt) in fixed_runtime_providers() {
            let consumer = || {
                let mut script = Chunk::new("<script>");
                crate::primitives::globals::emit_read(&mut script, receipt.global_name(), 0);
                script.emit_op(Op::RETURN, 0);
                vec![script]
            };
            let mut original = consumer();
            finalize_with_runtime_helpers(&mut original);
            let index = original.iter().position(|chunk| chunk.name == export).unwrap();
            let helper = &original[index];
            assert_eq!(helper.arity, receipt.arity());
            assert_eq!(helper.result_arity, receipt.result_arity());
            assert!(helper.types.is_empty());
            assert!(helper.direct_defined_calls.is_empty());
            let mut object = Object::new();
            object.kind = ObjectKind::Function(Function {
                name: None, arity: receipt.arity(), chunk_index: index, upvalues: Vec::new(),
            });
            let binding = Value::Object(vybe_runtime::heap::alloc(object));
            let mut vm = vybe_runtime::VM::new();
            vm.chunks = original;
            vm.set_global(receipt.global_name(), binding.clone());
            let installed = vm.snapshot_runtime_helper(receipt).unwrap();
            assert!(vm.validate_runtime_helper(&installed));

            let mut certificate = crate::compile_dependencies::Certificate::default();
            certificate.install_runtime_helpers(snapshot_runtime_helpers(&vm));
            let mut included = consumer();
            let excluded = certificate.runtime_helper_exports(&included);
            assert!(excluded.contains(&export));
            finalize_with_runtime_helpers_excluding(&mut included, &excluded);
            assert!(!included.iter().any(|chunk| chunk.name == export));
            assert!(included[0].global_inits.is_empty());

            // Equal labels, serialized metadata and a copied guest binding
            // cannot recreate the recipe identity or its installing VM scope.
            let forged = vybe_runtime::shared_provider::RuntimeHelperReceipt::new(
                receipt.global_name(), receipt.version(), receipt.arity(), receipt.result_arity());
            assert!(vm.snapshot_runtime_helper(&forged).is_none());
            let mut other = vybe_runtime::VM::new();
            other.chunks = vm.chunks.clone();
            other.set_global(receipt.global_name(), binding);
            assert!(!other.validate_runtime_helper(&installed));
            vm.set_global(receipt.global_name(), Value::Null);
            assert!(!certificate.runtime_helpers_valid(&vm));
            assert!(vm.snapshot_runtime_helper(receipt).is_none());
            let mut fallback = consumer();
            finalize_with_runtime_helpers(&mut fallback);
            assert!(fallback.iter().any(|chunk| chunk.name == export));
        }
    }

    fn comparison_consumer() -> Vec<Chunk> {
        let mut script = Chunk::new("<script>");
        crate::primitives::globals::emit_read(&mut script, "__vybe_dyneq", 0);
        script.emit_op(Op::RETURN, 0);
        vec![script]
    }

    #[test]
    fn runtime_helper_reuse_omits_only_attested_live_provider_bodies() {
        use vybe_runtime::value::{Function, Object, ObjectKind};
        let mut original = comparison_consumer();
        finalize_with_runtime_helpers(&mut original);
        let index = original.iter().position(|chunk| chunk.name == "__stdlib_dyneq").unwrap();
        assert!(original[index].runtime_helper_receipt.is_some());
        let mut vm = vybe_runtime::VM::new();
        vm.chunks = original;
        vm.set_global("__vybe_dyneq", Value::Object(vybe_runtime::heap::alloc(Object {
            properties: Default::default(), type_id: 0, descriptor: None, fields: Vec::new(),
            kind: ObjectKind::Function(Function {
                name: None, arity: 2, chunk_index: index, upvalues: Vec::new(),
            }),
        })));
        let mut certificate = crate::compile_dependencies::Certificate::default();
        certificate.install_runtime_helpers(snapshot_runtime_helpers(&vm));
        assert_eq!(certificate.runtime_helper_exports(&[]), ["__stdlib_dyneq"]);
        assert!(certificate.runtime_helpers_valid(&vm));
        let mut included = comparison_consumer();
        let excluded = certificate.runtime_helper_exports(&included);
        finalize_with_runtime_helpers_excluding(&mut included, &excluded);
        assert_eq!(included.len(), 1, "include must not construct another comparison body");
        assert!(included[0].global_inits.is_empty());
        assert!(referenced_helper_exports(&included).contains("__stdlib_dyneq"));

        let mut shadowed = comparison_consumer();
        shadowed[0].global_inits.push(vybe_runtime::chunk::GlobalInit {
            name: "__vybe_dyneq".into(), init: vybe_runtime::chunk::ConstExpr::RefFunc(0),
        });
        assert!(certificate.runtime_helper_exports(&shadowed).is_empty(),
            "local publication must not adopt a provider about to be replaced");

        // Ordinary compilation/export still owns a complete implementation.
        let mut standalone = comparison_consumer();
        finalize_with_runtime_helpers(&mut standalone);
        assert!(standalone.iter().any(|chunk| chunk.name == "__stdlib_dyneq"));
        vm.set_global("__vybe_dyneq", Value::Null);
        assert!(!certificate.runtime_helpers_valid(&vm));
        assert!(snapshot_runtime_helpers(&vm).is_empty());
    }

    #[test]
    fn runtime_helper_reuse_does_not_adopt_a_guest_function_name() {
        use vybe_runtime::value::{Function, Object, ObjectKind};
        let mut vm = vybe_runtime::VM::new();
        let mut impostor = Chunk::new("__stdlib_dyneq");
        impostor.arity = 2;
        impostor.local_count = 2;
        vm.chunks.push(impostor);
        vm.set_global("__vybe_dyneq", Value::Object(vybe_runtime::heap::alloc(Object {
            properties: Default::default(), type_id: 0, descriptor: None, fields: Vec::new(),
            kind: ObjectKind::Function(Function {
                name: Some("__stdlib_dyneq".into()), arity: 2, chunk_index: 0, upvalues: Vec::new(),
            }),
        })));
        assert!(snapshot_runtime_helpers(&vm).is_empty());
    }

    #[test]
    fn cached_helper_mappings_preserve_ordered_catalog_and_unknown_names() {
        let mut names: Vec<_> = mappings()
            .flat_map(|&(export, global)| [export, global])
            .collect();
        names.extend([
            "",
            "ordinary field",
            "__vybe_unknown",
            "__stdlib_unknown",
            "__ecma_regexp_unknown",
            "é😀",
        ]);
        for name in names {
            assert_eq!(
                helper_export_for_global(name),
                mappings().find_map(|&(export, global)| (name == global).then_some(export))
            );
            assert_eq!(
                helper_global_for_export(name),
                mappings().find_map(|&(export, global)| (name == export).then_some(global))
            );
        }
    }

    #[test]
    #[ignore = "interleaved opt-level-0 actual compiler lookup benchmark"]
    fn benchmark_actual_helper_mapping_lookup() {
        use std::{hint::black_box, time::Instant};
        let catalog: Vec<_> = mappings().copied().collect();
        let init = Instant::now();
        black_box(helper_mapping_indexes());
        println!(
            "mapping cache init={}ns entries={}",
            init.elapsed().as_nanos(),
            catalog.len()
        );
        for reverse in [false, true] {
            for case in ["miss", "all_hits", "mixed", "first_hit"] {
                let names: Vec<_> = (0..50_000)
                    .map(|i| {
                        let (export, global) = catalog[if case == "first_hit" {
                            0
                        } else {
                            i % catalog.len()
                        }];
                        if case == "miss" || (case == "mixed" && i % 10 != 0) {
                            format!("ordinary-property-{i}")
                        } else {
                            (if reverse { export } else { global }).to_owned()
                        }
                    })
                    .collect();
                let timed = |cached| {
                    let start = Instant::now();
                    let mut hits = 0usize;
                    for name in &names {
                        let name = black_box(name.as_str());
                        let value = if cached {
                            if reverse {
                                helper_global_for_export(name)
                            } else {
                                helper_export_for_global(name)
                            }
                        } else {
                            mappings().find_map(|&(export, global)| {
                                let (key, value) = if reverse {
                                    (export, global)
                                } else {
                                    (global, export)
                                };
                                (name == key).then_some(value)
                            })
                        };
                        hits += black_box(value).is_some() as usize;
                    }
                    (start.elapsed().as_nanos(), hits)
                };
                let (mut before, mut after) = (Vec::new(), Vec::new());
                for batch in 0..7 {
                    let (a, b) = if batch % 2 == 0 {
                        (timed(false), timed(true))
                    } else {
                        let b = timed(true);
                        let a = timed(false);
                        (a, b)
                    };
                    assert_eq!(a.1, b.1);
                    before.push(a.0);
                    after.push(b.0);
                }
                before.sort_unstable();
                after.sort_unstable();
                println!(
                    "actual mapping {case} reverse={reverse} old={}ns cached={}ns speedup={:.2}x",
                    before[3],
                    after[3],
                    before[3] as f64 / after[3] as f64
                );
            }
        }
    }
}
