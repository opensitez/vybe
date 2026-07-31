//! Kotlin's `common:kotlin.*` emit targets.
//!
//! The shared dispatcher hands a `common:<name>` target it does not recognise
//! to the language that declared it, which is how a language owns an emitter
//! without shared code learning its name.

use vybe_runtime::Chunk;

pub fn dispatch(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
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
        "kotlin.tostring" => {
            crate::emitter::tostring::emit_to_string(chunks, current, line);
            true
        }
        _ => false,
    }
}
