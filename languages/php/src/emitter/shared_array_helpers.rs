//! Share audited, caller-independent PHP adapter bodies within a module.
//! The normal dispatcher still owns argument order and the exact algorithm.

use vybe_runtime::chunk::{LocalName, LocalOrigin};
use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

const RECEIPT: &str = "<php-array-helper:v1:implode:boxed:caller-independent>";

// Array keys uses compiler-owned append-only unit indices, not the debug
// receipt used by the older implode wrapper below. Foreign Wasm/debug names
// cannot manufacture an Owner entry or an exact parser/producer registration.
fn array_keys_context_key(_: &[Chunk], _: u32) -> Vec<u8> {
    Vec::new()
}

fn build_array_keys_one(chunks: &mut Vec<Chunk>, _: u32) -> usize {
    let index = chunks.len();
    let mut chunk = Chunk::new("__stdlib_php_array_keys_1");
    chunk.arity = 1;
    chunk.local_count = 1;
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunks.push(chunk);
    // Keep the complete existing guard and algorithm, including its errors.
    assert!(super::dispatch_inline("php.array_keys", chunks, index, 1, 0));
    chunks[index].emit_op(Op::RETURN, 0);
    index
}

fn invoke_array_keys_one(chunks: &mut Vec<Chunk>, current: usize, helper: usize, line: u32) {
    assert!(vybe_compiler::primitives::functions::try_emit_direct_helper_call(
        chunks, current, helper, 1, line,
    ));
}

fn unused_array_keys_factory() -> Chunk {
    unreachable!("array_keys uses its unit-owned contextual factory")
}

fn unused_array_keys_caller(_: &mut Chunk, _: usize, _: u32) {
    unreachable!("array_keys uses its unit-owned contextual invocation")
}

fn array_keys_definition() -> vybe_compiler::primitives::helper_templates::Definition {
    use vybe_compiler::primitives::helper_templates::{Contract, Definition};
    Definition {
        contract: Contract::ContextualAppendOnlyUnit {
            key: array_keys_context_key,
            build: build_array_keys_one,
            invoke: invoke_array_keys_one,
        },
        operation: "php.array_keys", namespace: "php", version: 1,
        parser: crate::parse, producer: super::dispatch, arity: 1,
        helper_name: "__stdlib_php_array_keys_1",
        factory: unused_array_keys_factory, caller: unused_array_keys_caller,
    }
}

fn build_array_values_one(chunks: &mut Vec<Chunk>, _: u32) -> usize {
    let index = chunks.len();
    let mut chunk = Chunk::new("__stdlib_php_array_values_1");
    chunk.arity = 1;
    chunk.local_count = 1;
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunks.push(chunk);
    assert!(super::dispatch_inline("php.array_values", chunks, index, 1, 0));
    chunks[index].emit_op(Op::RETURN, 0);
    index
}

fn array_values_definition() -> vybe_compiler::primitives::helper_templates::Definition {
    use vybe_compiler::primitives::helper_templates::{Contract, Definition};
    Definition {
        contract: Contract::ContextualAppendOnlyUnit {
            key: array_keys_context_key,
            build: build_array_values_one,
            invoke: invoke_array_keys_one,
        },
        operation: "php.array_values", namespace: "php", version: 1,
        parser: crate::parse, producer: super::dispatch, arity: 1,
        helper_name: "__stdlib_php_array_values_1",
        factory: unused_array_keys_factory, caller: unused_array_keys_caller,
    }
}

fn build_array_comparison(chunks: &mut Vec<Chunk>, operation: &str, name: &str) -> usize {
    let index = chunks.len();
    let mut chunk = Chunk::new(name);
    chunk.arity = 2;
    chunk.local_count = 2;
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    chunks.push(chunk);
    assert!(super::dispatch_inline(operation, chunks, index, 2, 0));
    chunks[index].emit_op(Op::RETURN, 0);
    index
}

fn build_array_diff_two(chunks: &mut Vec<Chunk>, _: u32) -> usize {
    build_array_comparison(chunks, "php.array_diff", "__stdlib_php_array_diff_2")
}

fn build_array_intersect_two(chunks: &mut Vec<Chunk>, _: u32) -> usize {
    build_array_comparison(chunks, "php.array_intersect", "__stdlib_php_array_intersect_2")
}

fn invoke_array_comparison(chunks: &mut Vec<Chunk>, current: usize, helper: usize, line: u32) {
    assert!(vybe_compiler::primitives::functions::try_emit_direct_helper_call(
        chunks, current, helper, 2, line,
    ));
}

fn array_comparison_definition(intersect: bool) -> vybe_compiler::primitives::helper_templates::Definition {
    use vybe_compiler::primitives::helper_templates::{Contract, Definition};
    Definition {
        contract: Contract::ContextualAppendOnlyUnit {
            // Both bodies use only their parameters, fixed imports and private
            // snapshots. Runtime coercion/callbacks stay in the original body.
            key: array_keys_context_key,
            build: if intersect { build_array_intersect_two } else { build_array_diff_two },
            invoke: invoke_array_comparison,
        },
        operation: if intersect { "php.array_intersect" } else { "php.array_diff" },
        namespace: "php", version: 1,
        parser: crate::parse, producer: super::dispatch, arity: 2,
        helper_name: if intersect { "__stdlib_php_array_intersect_2" } else { "__stdlib_php_array_diff_2" },
        factory: unused_array_keys_factory, caller: unused_array_keys_caller,
    }
}

pub(crate) fn helper_definitions() -> &'static [vybe_compiler::primitives::helper_templates::Definition] {
    use vybe_compiler::primitives::helper_templates::Definition;
    static DEFINITIONS: std::sync::OnceLock<Vec<Definition>> = std::sync::OnceLock::new();
    DEFINITIONS.get_or_init(|| {
        // Preserve all existing serializer recipes and their exact callbacks.
        let mut definitions: Vec<_> = super::super::serialization_adapter::HELPER_DEFINITIONS
            .iter().map(|definition| Definition {
                contract: definition.contract,
                operation: definition.operation, namespace: definition.namespace,
                version: definition.version, parser: definition.parser,
                producer: definition.producer, arity: definition.arity,
                helper_name: definition.helper_name,
                factory: definition.factory, caller: definition.caller,
            }).collect();
        definitions.push(super::super::copy_adapter::strict_eq_definition());
        definitions.push(array_keys_definition());
        definitions.push(array_comparison_definition(false));
        definitions.push(array_comparison_definition(true));
        definitions.push(array_values_definition());
        definitions.push(super::super::string_adapter::trim_one_definition());
        definitions.push(super::super::string_adapter::quotemeta_string_definition());
        definitions.push(super::super::string_adapter::count_chars_definition(1));
        definitions.push(super::super::string_adapter::count_chars_definition(2));
        definitions
    })
}

pub(super) fn emit(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    // Do not infer shareability from an adapter name or signature alone.
    // Implode has no caller locals, reference writes or lexical receiver use.
    // Other adapters retain their original emissions until separately audited.
    if name != "php.implode" || !matches!(argc, 1 | 2) {
        return false;
    }
    let helper_name = format!("__stdlib_php_implode_{argc}");
    let helper = chunks.iter().position(|chunk| {
        chunk.name == helper_name
            && chunk.source_path.is_none()
            && chunk.arity == argc
            && chunk.result_arity == 1
            && chunk.func_sig.is_none()
            && !chunk.takes_receiver
            && !chunk.is_async
            && !chunk.is_generator
            && chunk.capture_count == 0
            && chunk.direct_entry_has_captures != Some(true)
            && chunk.local_names.iter().any(|local| {
                local.slot == 0
                    && local.origin == LocalOrigin::Compiler
                    && local.name == RECEIPT
            })
    });
    let helper = helper.unwrap_or_else(|| {
        let mut chunk = Chunk::new(helper_name);
        chunk.arity = argc;
        chunk.local_count = u16::from(argc);
        chunk.local_names.push(LocalName::new(0, RECEIPT, LocalOrigin::Compiler));
        for slot in 0..u16::from(argc) {
            chunk.emit_op_u16(Op::LOCAL_GET, slot, 0);
        }
        let index = chunks.len();
        chunks.push(chunk);
        // Bypass the sharing wrapper while emitting the original body.
        assert!(super::dispatch_inline(name, chunks, index, argc, 0));
        chunks[index].emit_op(Op::RETURN, 0);
        index
    });
    // Failed eligibility leaves the caller untouched, so the normal inline
    // dispatcher remains a semantics-preserving fallback.
    vybe_compiler::primitives::functions::try_emit_direct_helper_call(
        chunks, current, helper, argc, line,
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn repeated_implodes_share_the_original_loop_body() {
        for argc in [1, 2] {
            let mut shared = vec![Chunk::new("caller".to_string())];
            let mut inline = vec![Chunk::new("caller".to_string())];
            for _ in 0..20 {
                assert!(emit("php.implode", &mut shared, 0, argc, 7));
                assert!(super::super::dispatch_inline(
                    "php.implode", &mut inline, 0, argc, 7,
                ));
            }
            assert_eq!(shared.len(), 2);
            assert_eq!(shared[0].local_count, 0);
            let shared_bytes: usize = shared.iter().map(|chunk| chunk.code.len()).sum();
            assert!(shared_bytes * 5 < inline[0].code.len());
        }
    }

    #[test]
    fn arities_do_not_share_incompatible_bodies() {
        let mut chunks = vec![Chunk::new("caller".to_string())];
        assert!(emit("php.implode", &mut chunks, 0, 1, 7));
        assert!(emit("php.implode", &mut chunks, 0, 2, 7));
        assert_eq!(chunks.len(), 3);
        assert_eq!(chunks[1].arity, 1);
        assert_eq!(chunks[2].arity, 2);
    }

    #[test]
    fn source_names_cannot_impersonate_the_private_body() {
        let mut source = Chunk::new("__stdlib_php_implode_1".to_string());
        source.arity = 1;
        let mut chunks = vec![Chunk::new("caller".to_string()), source];
        assert!(emit("php.implode", &mut chunks, 0, 1, 7));
        assert_eq!(chunks.len(), 3);
    }

    #[test]
    fn unaudited_operations_and_arities_are_untouched() {
        let mut chunks = vec![Chunk::new("caller".to_string())];
        for (name, argc) in [("php.in_array", 2), ("php.implode", 0), ("php.implode", 3)] {
            assert!(!emit(name, &mut chunks, 0, argc, 7));
        }
        assert_eq!(chunks.len(), 1);
        assert!(chunks[0].code.is_empty());
    }
}
