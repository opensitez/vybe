//! Auto-extracted `cobol.*` dispatch (language-specific routing lives in the
//! language module; the common dispatcher delegates here).

use vybe_bytecode::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        "cobol.round_away_from_zero" => {
            crate::emitter::arithmetic::emit_round_away_from_zero(chunks, current, line)
        }
        "cobol.integer_of_date" => {
            crate::emitter::date::emit_integer_of_date(chunks, current, line)
        }
        "cobol.cancel" => crate::emitter::control::emit_cancel(chunks, current, argc, line),
        "cobol.release" => crate::emitter::files::emit_release(chunks, current, argc, line),
        "cobol.return_record" => {
            crate::emitter::files::emit_return_record(chunks, current, argc, line)
        }
        "cobol.alter" => crate::emitter::control::emit_alter(chunks, current, argc, line),
        "cobol.sort" => crate::emitter::files::emit_sort(chunks, current, argc, line),
        "cobol.merge" => crate::emitter::files::emit_merge(chunks, current, argc, line),
        "cobol.copy" => crate::emitter::data::emit_copy(chunks, current, argc, line),
        "cobol.validate" => crate::emitter::data::emit_validate(chunks, current, argc, line),
        "cobol.typedef" => crate::emitter::data::emit_typedef(chunks, current, argc, line),
        "cobol.move_corresponding" => {
            crate::emitter::data::emit_move_corresponding(chunks, current, argc, line)
        }

        // ── String ops ──
        _ => return false,
    }
    true
}
