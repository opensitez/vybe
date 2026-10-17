//! PHP stdout / output-buffer helpers.
//!
//! PHP's stringification of a value, then the SHARED writer. The buffer stack
//! itself lives in `vybe_compiler::primitives::io` — it is not a PHP concept,
//! and keeping a private copy here is what let `var_dump` escape `ob_start()`.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

/// Stringify one value PHP's way, then hand it to the SHARED writer.
///
/// The buffer check is deliberately not repeated here: `emit_write_or_buffer`
/// owns it, so `echo`, `print` and `var_dump` cannot drift apart on whether
/// output gets captured — which is precisely how `var_dump` used to escape
/// `ob_start()`.
const ECHO_WRITE_CHUNK: &str = "__stdlib_php_echo_write_one";
const VAR_DUMP_WRITE_CHUNK: &str = "__stdlib_php_var_dump_write_one";
const STRING_WRITE_CHUNK: &str = "__stdlib_php_string_write_one";
const STRING_WRITE_RECEIPT: &str = "<php-string-output:v1:line0>";

/// A physically proven String needs buffering/output, but no PHP coercion.
/// Share the writer body instead of expanding buffer checks at each literal.
pub fn emit_php_string_output(chunks: &mut Vec<Chunk>, current: usize, print: bool, line: u32) {
    use vybe_runtime::chunk::{LocalName, LocalOrigin};
    let helper = if let Some(index) = chunks.iter().position(|chunk|
        chunk.name == STRING_WRITE_CHUNK && chunk.source_path.is_none()
        && chunk.arity == 1 && chunk.result_arity == 1 && chunk.func_sig.is_none()
        && !chunk.takes_receiver && !chunk.is_async && !chunk.is_generator
        && chunk.capture_count == 0 && chunk.direct_entry_has_captures != Some(true)
        && chunk.local_names.iter().any(|local| local.slot == 0
            && local.origin == LocalOrigin::Compiler && local.name == STRING_WRITE_RECEIPT))
    {
        index
    } else {
        let index = chunks.len();
        let mut chunk = Chunk::new(STRING_WRITE_CHUNK);
        chunk.arity = 1;
        chunk.local_count = 1;
        chunk.local_names.push(LocalName::new(0, STRING_WRITE_RECEIPT, LocalOrigin::Compiler));
        chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
        chunks.push(chunk);
        vybe_compiler::primitives::io::emit_write_or_buffer(chunks, index, 0);
        chunks[index].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunks[index].emit_op(Op::RETURN, 0);
        index
    };
    call_void_helper_from_stack(chunks, current, helper, line);
    if print {
        chunks[current].emit_i32_const(1, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
}

fn ensure_echo_write_chunk(chunks: &mut Vec<Chunk>) -> usize {
    if let Some(index) = chunks
        .iter()
        .position(|chunk| chunk.name == ECHO_WRITE_CHUNK)
    {
        return index;
    }

    let index = chunks.len();
    let mut helper = Chunk::new(ECHO_WRITE_CHUNK);
    helper.arity = 1;
    helper.local_count = 1;
    chunks.push(helper);
    chunks[index].emit_op_u16(Op::LOCAL_GET, 0, 0);
    emit_php_stdout_write(chunks, index, 0);
    chunks[index].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunks[index].emit_op(Op::RETURN, 0);
    index
}

fn ensure_var_dump_write_chunk(chunks: &mut Vec<Chunk>) -> usize {
    if let Some(index) = chunks
        .iter()
        .position(|chunk| chunk.name == VAR_DUMP_WRITE_CHUNK)
    {
        return index;
    }

    let index = chunks.len();
    let mut helper = Chunk::new(VAR_DUMP_WRITE_CHUNK);
    helper.arity = 1;
    helper.local_count = 1;
    chunks.push(helper);
    chunks[index].emit_op_u16(Op::LOCAL_GET, 0, 0);
    super::string_adapter::emit_var_dump_stringify(chunks, index, 1, 0);
    chunks[index].emit_string_const("\n", 0);
    vybe_compiler::primitives::strings::emit_concat(&mut chunks[index], 2, 0);
    vybe_compiler::primitives::io::emit_write_or_buffer(chunks, index, 0);
    chunks[index].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunks[index].emit_op(Op::RETURN, 0);
    index
}

fn call_void_helper_from_stack(
    chunks: &mut [Chunk],
    current: usize,
    helper: usize,
    line: u32,
) {
    if vybe_compiler::primitives::functions::try_emit_direct_helper_call(
        chunks, current, helper, 1, line,
    ) {
        chunks[current].emit_op(Op::DROP, line);
        return;
    }
    let val_slot = alloc_local(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, val_slot, line);
    chunks[current].emit_op_u16(Op::REF_FUNC, helper as u16, line);
    chunks[current].emit(0u8, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, val_slot, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(&mut chunks[current], 1, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn write_helper_arguments(
    chunks: &mut [Chunk],
    current: usize,
    helper: usize,
    argc: u8,
    line: u32,
) {
    if argc == 0 {
        return;
    }
    // Keep the first argument on the stack. Only later arguments must be
    // parked while earlier values are stringified/written in source order.
    let count = u16::from(argc - 1);
    let base = chunks[current].alloc_scratch(count);
    for offset in (0..count).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    call_void_helper_from_stack(chunks, current, helper, line);
    for offset in 0..count {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + offset, line);
        call_void_helper_from_stack(chunks, current, helper, line);
    }
}

pub fn emit_php_stdout_write(chunks: &mut [Chunk], current: usize, line: u32) {
    super::string_adapter::emit_echo_stringify(chunks, current, 1, line);
    vybe_compiler::primitives::io::emit_write_or_buffer(chunks, current, line);
}

pub fn emit_php_echo(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 0 {
        let helper = ensure_echo_write_chunk(chunks);
        write_helper_arguments(chunks, current, helper, argc, line);
    }
    // The PHP walker emits echo as an expression-statement call. Its
    // statement owns the final DROP; the per-value helper result was already
    // discarded above, so retain the adapter's expression result here.
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `var_dump($a, $b, …)` — one type-annotated dump record per argument.
///
/// Same argc-driven spill as [`emit_php_echo`]: the arguments are already on
/// the stack, so they come off in reverse and are replayed in source order.
/// Like `echo`, it goes through the shared writer, so `ob_start(); var_dump($x);`
/// captures. It previously wrote straight to `wasi:logging` and escaped every
/// active buffer.
pub fn emit_php_var_dump(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let helper = ensure_var_dump_write_chunk(chunks);
    write_helper_arguments(chunks, current, helper, argc, line);
}

pub fn emit_php_print_expr(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let helper = ensure_echo_write_chunk(chunks);
    call_void_helper_from_stack(chunks, current, helper, line);
    chunks[current].emit_i32_const(1, line);
}

pub fn emit_ob_implicit_flush(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    chunks[current].emit_bool_const(true, line);
}

pub fn emit_ob_list_handlers(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    vybe_compiler::primitives::io::emit_ob_list_handlers(
        chunks,
        current,
        "default output handler",
        line,
    );
}

pub fn emit_ob_gzhandler(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    chunks[current].emit_bool_const(false, line);
}

/// PHP measures buffered output in UTF-8 BYTES (`strlen`), not code units, so
/// `ob_get_length()` on "éclair" is 7 and not 6. Reuses the same emitter
/// `strlen` itself uses rather than spelling the UTF-8 walk a second time.
fn php_buffer_len(chunks: &mut [Chunk], current: usize, line: u32) {
    super::string_adapter::emit_strlen(chunks, current, 1, line);
}

/// `ob_get_length()` — PHP's byte count, not the shared code-unit count.
pub fn emit_ob_get_length(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    vybe_compiler::primitives::io::emit_ob_get_length(chunks, current, Some(php_buffer_len), line);
}

/// `ob_get_status()` — PHP supplies its own word for a frame with no handler
/// and its own byte count; the shared primitive has neither.
pub fn emit_ob_get_status(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    vybe_compiler::primitives::io::emit_ob_get_status(
        chunks,
        current,
        argc,
        "default output handler",
        Some(php_buffer_len),
        line,
    );
}
