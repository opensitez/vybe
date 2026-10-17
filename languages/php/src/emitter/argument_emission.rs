//! Optional physical-operand specializations of existing PHP emitters.
//! Selection and producer authority belong to the generic compiler contract;
//! this callback does not resolve source names or infer variable types.

use vybe_compiler::primitives::argument_emission::{Definition, ProducedValueProof};
use vybe_runtime::{Chunk, Op};
use vybe_runtime::registry::emitter_results::{EmitOutcome, EmitRequest, ResultRepresentation};

fn emit_string_coercion_cold(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let helper = super::string_adapter::ensure_concat_stringify_helper(chunks);
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_TEE, value, line);
    let test = chunks[current].add_import("wasm:js-string", "test");
    chunks[current].emit_call(test, 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    assert!(vybe_compiler::primitives::functions::try_emit_direct_helper_call(
        chunks, current, helper, 1, line,
    ), "owned PHP stringify helper must retain its capture-free generic ABI");
    chunks[current].emit_end(line);
}

pub fn emit_common_with_operand_facts(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
    request: EmitRequest,
    operands: &[ProducedValueProof],
) -> EmitOutcome {
    // A physically produced immutable string is already in the scanner's ABI.
    // Do not inspect its constant/import spelling or move unknown-value PHP
    // conversion (including object callbacks) into a different call frame.
    if name == "php.trim" && argc == 1 && operands.len() == 1
        && matches!(request, EmitRequest::Value | EmitRequest::DiscardedValue)
        && matches!(operands[0], ProducedValueProof::StringLiteral { .. })
    {
        vybe_compiler::primitives::strings::emit_trim_chars(
            chunks, current, 1,
            vybe_compiler::primitives::strings::TrimOptions::both(Some(" \t\n\r\0\x0B")),
            line,
        );
        return EmitOutcome::Handled(ResultRepresentation::Unknown);
    }
    if matches!(request, EmitRequest::Value | EmitRequest::DiscardedValue)
        && argc == 1
        && matches!(operands, [ProducedValueProof::StringLiteral { .. }])
        && matches!(name, "php.urlencode" | "php.rawurlencode")
    {
        use vybe_compiler::primitives::url;
        let options = if name == "php.urlencode" {
            url::PercentOptions::form()
        } else {
            url::PercentOptions::rfc3986()
        };
        url::emit_percent_encode(chunks, current, options, line);
        return EmitOutcome::Handled(ResultRepresentation::Unknown);
    }
    if name == "php.print_r" && matches!(argc, 1 | 2)
        && operands.len() == usize::from(argc)
        && matches!(request, EmitRequest::Value | EmitRequest::DiscardedValue)
        && matches!(operands[0], ProducedValueProof::StringLiteral { .. })
    {
        let value = if argc == 2 {
            let flag = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, flag, line);
            let value = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
            super::array_adapter::emit_php_empty_from_slot(chunks, current, flag, line);
            chunks[current].emit_op(Op::I32_EQZ, line);
            chunks[current].emit_if_value(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
            Some(value)
        } else { None };
        super::output_adapter::emit_php_string_output(chunks, current, false, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_bool_const(true, line);
        if value.is_some() { chunks[current].emit_end(line); }
        return EmitOutcome::Handled(ResultRepresentation::Unknown);
    }
    if argc == 1 && operands.len() == 1
        && matches!(request, EmitRequest::Value | EmitRequest::DiscardedValue)
        && matches!(name, "php.echo" | "php.print_expr")
        && matches!(operands[0], ProducedValueProof::StringLiteral { .. })
    {
        super::output_adapter::emit_php_string_output(chunks, current, name == "php.print_expr", line);
        return EmitOutcome::Handled(ResultRepresentation::Unknown);
    }
    // Physical strings need no conversion. Keep their common path inline;
    // share the full PHP conversion ladder only on the nonstring branch.
    if matches!(request, EmitRequest::Value | EmitRequest::DiscardedValue)
        && operands.len() == usize::from(argc)
        && matches!((name, argc), ("php.strrev", 1) | ("php.str_repeat", 2)
            | ("php.strtolower", 1) | ("php.strtoupper", 1))
    {
        if !matches!(operands.first(), Some(ProducedValueProof::StringLiteral { .. })) {
            let count = if argc == 2 {
                let slot = chunks[current].alloc_scratch(1);
                chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
                Some(slot)
            } else { None };
            emit_string_coercion_cold(chunks, current, line);
            if let Some(slot) = count {
                chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
            }
        }
        let handled = match (name, argc) {
            ("php.strrev", 1) => {
                vybe_compiler::primitives::strings::emit_str_reverse(&mut chunks[current], line);
                true
            }
            ("php.str_repeat", 2) => {
                vybe_compiler::primitives::strings::emit_repeat(&mut chunks[current], line);
                true
            }
            ("php.strtolower", 1) => {
                vybe_compiler::primitives::strings::emit_to_lower(&mut chunks[current], line);
                true
            }
            ("php.strtoupper", 1) => {
                vybe_compiler::primitives::strings::emit_to_upper(&mut chunks[current], line);
                true
            }
            _ => false,
        };
        if handled {
            return EmitOutcome::Handled(ResultRepresentation::Unknown);
        }
    }
    // The standard PHP native provider guarantees a wire-format String for
    // every immutable String operand. Unknown values retain the portable arm.
    if name == "php.serialize"
        && argc == 1
        && operands.len() == 1
        && matches!(request, EmitRequest::Value | EmitRequest::DiscardedValue)
        && matches!(operands[0], ProducedValueProof::StringLiteral { .. })
    {
        let chunk = &mut chunks[current];
        let encode = chunk.add_import("php:serialization", "encodeScalarArray");
        chunk.emit_call(encode, 1, line);
        return EmitOutcome::Handled(ResultRepresentation::Unknown);
    }
    if !matches!(request, EmitRequest::Value | EmitRequest::DiscardedValue)
        || name != "php.str_replace"
        || !matches!(argc, 3 | 4)
        || operands.len() != usize::from(argc)
        || !operands[..2]
            .iter()
            .all(|p| matches!(p, ProducedValueProof::StringLiteral { .. }))
    {
        return EmitOutcome::Unhandled;
    }
    if matches!(operands[2], ProducedValueProof::StringLiteral { .. }) {
        super::string_adapter::emit_str_replace_proven_strings(chunks, current, argc, line);
    } else {
        super::string_adapter::emit_str_replace_proven_search_replace_strings(
            chunks, current, argc, line,
        );
    }
    // The string result retains the ordinary PHP value-position ABI, including
    // subsequent language-owned truthiness or result-discard handling.
    EmitOutcome::Handled(ResultRepresentation::Unknown)
}

/// Native serialization and PHP string operations with shared cold coercion.
/// Replacing the PHP serialization provider must preserve its guarantee that
/// immutable String input returns a serialized String, never the null sentinel.
pub static ARGUMENT_DEFINITIONS: &[Definition] = &[
    Definition { operation: "php.trim", parser: crate::parse,
        producer: super::dispatch::dispatch, dispatch: emit_common_with_operand_facts },
    Definition { operation: "php.urlencode", parser: crate::parse,
        producer: super::dispatch::dispatch, dispatch: emit_common_with_operand_facts },
    Definition { operation: "php.rawurlencode", parser: crate::parse,
        producer: super::dispatch::dispatch, dispatch: emit_common_with_operand_facts },
    Definition { operation: "php.print_r", parser: crate::parse,
        producer: super::dispatch::dispatch, dispatch: emit_common_with_operand_facts },
    Definition { operation: "php.echo", parser: crate::parse,
        producer: super::dispatch::dispatch, dispatch: emit_common_with_operand_facts },
    Definition { operation: "php.print_expr", parser: crate::parse,
        producer: super::dispatch::dispatch, dispatch: emit_common_with_operand_facts },
    Definition { operation: "php.serialize", parser: crate::parse,
        producer: super::dispatch::dispatch, dispatch: emit_common_with_operand_facts },
    Definition { operation: "php.strrev", parser: crate::parse,
        producer: super::dispatch::dispatch, dispatch: emit_common_with_operand_facts },
    Definition { operation: "php.str_repeat", parser: crate::parse,
        producer: super::dispatch::dispatch, dispatch: emit_common_with_operand_facts },
    Definition { operation: "php.strtolower", parser: crate::parse,
        producer: super::dispatch::dispatch, dispatch: emit_common_with_operand_facts },
    Definition { operation: "php.strtoupper", parser: crate::parse,
        producer: super::dispatch::dispatch, dispatch: emit_common_with_operand_facts },
];
