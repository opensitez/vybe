//! Auto-extracted `vb.*` dispatch (language-specific routing lives in the
//! language module; the common dispatcher delegates here).

use vybe_bytecode::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        "vb.pmt" => crate::emitter::vb::financial_adapter::emit_vb_pmt(chunks, current, argc, line),
        "vb.fv" => crate::emitter::vb::financial_adapter::emit_vb_fv(chunks, current, argc, line),
        "vb.pv" => crate::emitter::vb::financial_adapter::emit_vb_pv(chunks, current, argc, line),
        "vb.nper" => {
            crate::emitter::vb::financial_adapter::emit_vb_nper(chunks, current, argc, line)
        }
        "vb.rate" => {
            crate::emitter::vb::financial_adapter::emit_vb_rate(chunks, current, argc, line)
        }
        "vb.ipmt" => {
            crate::emitter::vb::financial_adapter::emit_vb_ipmt(chunks, current, argc, line)
        }
        "vb.ppmt" => {
            crate::emitter::vb::financial_adapter::emit_vb_ppmt(chunks, current, argc, line)
        }
        "vb.sln" => crate::emitter::vb::financial_adapter::emit_vb_sln(chunks, current, argc, line),
        "vb.ddb" => crate::emitter::vb::financial_adapter::emit_vb_ddb(chunks, current, argc, line),
        "vb.syd" => crate::emitter::vb::financial_adapter::emit_vb_syd(chunks, current, argc, line),
        "vb.dir" => crate::emitter::vb::misc_adapter::emit_vb_dir(chunks, current, argc, line),
        "vb.filedatetime" => {
            crate::emitter::vb::misc_adapter::emit_vb_filedatetime(chunks, current, argc, line)
        }
        "vb.lof" => crate::emitter::vb::misc_adapter::emit_vb_lof(chunks, current, argc, line),
        "vb.eof" => crate::emitter::vb::misc_adapter::emit_vb_eof(chunks, current, argc, line),
        "vb.shell_pid" => {
            crate::emitter::vb::misc_adapter::emit_vb_shell_pid(chunks, current, argc, line)
        }
        _ => return false,
    }
    true
}
