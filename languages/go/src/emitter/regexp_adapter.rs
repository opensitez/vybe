use vybe_compiler::primitives::instructions::host;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

pub fn emit_helper(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    match name {
        "go.regex_split_pat_first" => {
            let base = alloc_locals(&mut chunks[current], 2);
            local_set(&mut chunks[current], base + 1, line);
            local_set(&mut chunks[current], base, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
            host::emit(&mut chunks[current], "ecma:regexp", "split", 2, line);
        }
        _ => {
            let _ = argc;
            return false;
        }
    }
    true
}

fn alloc_locals(chunk: &mut Chunk, count: u16) -> u16 {
    chunk.alloc_scratch(count)
}

fn local_set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}
