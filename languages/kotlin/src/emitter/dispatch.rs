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
        "kotlin.cmp_lt0" => {
            crate::emitter::numbers::emit_compare_zero(chunks, current, crate::emitter::numbers::CompareZero::Lt, line);
            true
        }
        "kotlin.cmp_gt0" => {
            crate::emitter::numbers::emit_compare_zero(chunks, current, crate::emitter::numbers::CompareZero::Gt, line);
            true
        }
        "kotlin.cmp_le0" => {
            crate::emitter::numbers::emit_compare_zero(chunks, current, crate::emitter::numbers::CompareZero::Le, line);
            true
        }
        "kotlin.cmp_ge0" => {
            crate::emitter::numbers::emit_compare_zero(chunks, current, crate::emitter::numbers::CompareZero::Ge, line);
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
        _ => false,
    }
}
