//! Calls made only after the owning PHP foreach lowering proves Iterator/Generator.
use vybe_runtime::{Chunk, Op};

fn emit_generator_operation(chunks: &mut [Chunk], current: usize, method: &str, line: u32) {
    match method {
        "rewind" => super::array_adapter::emit_generator_rewind(chunks, current, 1, line),
        "valid" => super::array_adapter::emit_generator_valid(chunks, current, 1, line),
        "current" => super::array_adapter::emit_generator_current(chunks, current, 1, line),
        "next" => super::array_adapter::emit_generator_next(chunks, current, 1, line),
        "key" => super::array_adapter::emit_generator_key(chunks, current, 1, line),
        _ => unreachable!(),
    }
}

fn ensure_generator_operation(chunks: &mut Vec<Chunk>, method: &str) -> usize {
    use vybe_runtime::chunk::{LocalName, LocalOrigin};

    let name = format!("__stdlib_php_foreach_generator_{method}");
    let receipt = format!("<php-foreach-generator:v1:{method}>");
    if let Some(index) = chunks.iter().position(|chunk|
        chunk.name == name && chunk.source_path.is_none()
        && chunk.arity == 1 && chunk.result_arity == 1 && chunk.func_sig.is_none()
        && !chunk.takes_receiver && !chunk.is_async && !chunk.is_generator
        && chunk.capture_count == 0 && chunk.direct_entry_has_captures != Some(true)
        && chunk.local_names.iter().any(|local| local.slot == 0
            && local.origin == LocalOrigin::Compiler && local.name == receipt))
    {
        return index;
    }

    let index = chunks.len();
    let mut helper = Chunk::new(name);
    helper.arity = 1;
    helper.local_count = 1;
    helper.local_names.push(LocalName::new(0, receipt, LocalOrigin::Compiler));
    helper.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunks.push(helper);
    emit_generator_operation(chunks, index, method, 0);
    chunks[index].emit_op(Op::RETURN, 0);
    index
}

pub fn emit(chunks: &mut Vec<Chunk>, current: usize, method: &str, line: u32) {
    let generator_helper = (method != "getiterator")
        .then(|| ensure_generator_operation(chunks, method));
    let receiver = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, receiver, line);
    let branch = if method != "getiterator" {
        let c = &mut chunks[current];
        c.emit_op_u16(Op::LOCAL_GET, receiver, line);
        let test = c.add_import("ecma:value", "isGenerator");
        c.emit_call(test, 1, line);
        let cast = c.add_import("wasm:js-boolean", "cast");
        c.emit_call(cast, 1, line);
        let branch = c.emit_if_value(line);
        c.emit_op_u16(Op::LOCAL_GET, receiver, line);
        // The direct-call linker owns function-index limits. Keep the inline
        // operation as its fallback instead of truncating a wide ref.func index.
        if !vybe_compiler::primitives::functions::try_emit_direct_helper_call(
            chunks, current, generator_helper.unwrap(), 1, line,
        ) {
            emit_generator_operation(chunks, current, method, line);
        }
        chunks[current].emit_else(line);
        Some(branch)
    } else {
        None
    };
    let c = &mut chunks[current];
    let slot = vybe_compiler::primitives::class_slots::resolve(
        &vybe_compiler::primitives::class_slots::ClassSlot::internal(if method == "getiterator" {
            "getIterator"
        } else {
            method
        }),
        &vybe_compiler::primitives::class_slots::PlainNames,
    );
    vybe_compiler::primitives::class_slots::emit_class_get(
        c,
        vybe_compiler::primitives::class_slots::ObjSource::Local(receiver),
        &slot,
        vybe_compiler::primitives::class_slots::Dest::Stack,
        line,
    );
    c.emit_op_u16(Op::LOCAL_GET, receiver, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(c, 1, line);
    if let Some(branch) = branch {
        c.emit_end(line);
        c.patch_block(branch);
    }
}
