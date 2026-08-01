//! Kotlin's `common:kotlin.*` emit targets.
//!
//! The shared dispatcher hands a `common:<name>` target it does not recognise
//! to the language that declared it, which is how a language owns an emitter
//! without shared code learning its name.

use vybe_runtime::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    if name.starts_with("java.") {
        return vybe_language_java::emitter::dispatch::dispatch(name, chunks, current, argc, line);
    }
    match name {
        "kotlin.print" => {
            crate::emitter::tostring::emit_print(chunks, current, argc, line);
            true
        }
        "kotlin.print_double" => {
            crate::emitter::numbers::emit_print_double(chunks, current, argc, line);
            true
        }
        "kotlin.is_infinite" => {
            crate::emitter::numbers::emit_is_infinite(chunks, current, argc, line);
            true
        }
        "kotlin.to_int_or_null" => {
            crate::emitter::numbers::emit_to_int_or_null(chunks, current, argc, line);
            true
        }
        "kotlin.to_double_or_null" => {
            crate::emitter::numbers::emit_to_double_or_null(chunks, current, argc, line);
            true
        }
        "kotlin.cmp_lt0" => {
            crate::emitter::numbers::emit_compare_zero(
                chunks,
                current,
                crate::emitter::numbers::CompareZero::Lt,
                line,
            );
            true
        }
        "kotlin.cmp_gt0" => {
            crate::emitter::numbers::emit_compare_zero(
                chunks,
                current,
                crate::emitter::numbers::CompareZero::Gt,
                line,
            );
            true
        }
        "kotlin.cmp_le0" => {
            crate::emitter::numbers::emit_compare_zero(
                chunks,
                current,
                crate::emitter::numbers::CompareZero::Le,
                line,
            );
            true
        }
        "kotlin.cmp_ge0" => {
            crate::emitter::numbers::emit_compare_zero(
                chunks,
                current,
                crate::emitter::numbers::CompareZero::Ge,
                line,
            );
            true
        }
        "kotlin.tostring" => {
            crate::emitter::tostring::emit_to_string(chunks, current, line);
            true
        }
        "kotlin.join_to_string" => {
            crate::emitter::tostring::emit_join_to_string(chunks, current, argc, line);
            true
        }
        "kotlin.not_null_assert" => {
            crate::emitter::nullability::emit_not_null_assert(chunks, current, argc, line);
            true
        }
        "kotlin.class_of" => {
            crate::emitter::nullability::emit_class_of(chunks, current, argc, line);
            true
        }
        "kotlin.error" => {
            crate::emitter::nullability::emit_error(chunks, current, argc, line);
            true
        }
        "kotlin.exception" => {
            crate::emitter::nullability::emit_exception(chunks, current, argc, "Exception", line);
            true
        }
        "kotlin.illegal_argument_exception" => {
            crate::emitter::nullability::emit_exception(
                chunks,
                current,
                argc,
                "IllegalArgumentException",
                line,
            );
            true
        }
        "kotlin.illegal_state_exception" => {
            crate::emitter::nullability::emit_exception(
                chunks,
                current,
                argc,
                "IllegalStateException",
                line,
            );
            true
        }
        "kotlin.null_pointer_exception" => {
            crate::emitter::nullability::emit_exception(
                chunks,
                current,
                argc,
                "NullPointerException",
                line,
            );
            true
        }
        "kotlin.index_out_of_bounds_exception" => {
            crate::emitter::nullability::emit_exception(
                chunks,
                current,
                argc,
                "IndexOutOfBoundsException",
                line,
            );
            true
        }
        "kotlin.require" => {
            crate::emitter::nullability::emit_precondition(
                chunks,
                current,
                argc,
                "IllegalArgumentException",
                line,
            );
            true
        }
        "kotlin.check" => {
            crate::emitter::nullability::emit_precondition(
                chunks,
                current,
                argc,
                "IllegalStateException",
                line,
            );
            true
        }
        "kotlin.require_not_null" => {
            crate::emitter::nullability::emit_precondition_not_null(
                chunks,
                current,
                argc,
                "IllegalArgumentException",
                line,
            );
            true
        }
        "kotlin.check_not_null" => {
            crate::emitter::nullability::emit_precondition_not_null(
                chunks,
                current,
                argc,
                "IllegalStateException",
                line,
            );
            true
        }
        _ => false,
    }
}
