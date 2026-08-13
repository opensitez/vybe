//! Auto-extracted `dotnet.*` dispatch (language-specific routing lives in the
//! language module; the common dispatcher delegates here).

use std::sync::Arc;
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::{Chunk, Value};

use vybe_runtime::opcode::Op;

/// `Control.CreateGraphics()` — construct a `Graphics` (via its registered
/// class global) stamped with the receiver control's `__control_name`, so
/// drawings route to the control's canvas. This mirrors the former
/// `CONTROL_CREATE_GRAPHICS` method Body, but is emitted at the call site
/// through the component descriptor (`MethodBody::Common`) rather than a
/// ctor-bound thunk: control leaves no longer emit a per-class ctor chunk,
/// so a host-constructed control must resolve `CreateGraphics` here.
///
/// Stack on entry: `[control]`; on exit: `[graphics]`.
fn emit_control_create_graphics(chunk: &mut Chunk, line: u32) {
    let name_key = chunk.add_constant(Value::String(Arc::from("__control_name")));
    // Construct Graphics via its host factory (its descriptor constructor
    // backing) rather than a class global — Graphics no longer emits a ctor
    // global now that its methods resolve through the descriptor.
    let graphics_new = chunk.add_import("vybe:gui", "graphicsNew");
    let name_slot = chunk.alloc_scratch(1);
    // Stash the control's name (consuming the control), then build a fresh
    // Graphics and copy the name onto it.
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, name_key, line); // [name]
    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line); // []
    chunk.emit_call(graphics_new, 0, line); // [graphics]
    core_wasm::dup(chunk, line); // [graphics, graphics]
    chunk.emit_op_u16(Op::LOCAL_GET, name_slot, line); // [graphics, graphics, name]
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, name_key, line); // [graphics, graphics]
    chunk.emit_op(Op::DROP, line); // [graphics]
}

/// VB `Choose(idx, v1, v2, ..., vN)` — variadic 1-indexed selector.
/// Packs the trailing values into an array, then `ARRAY_GET array[idx-1]`.
fn emit_choose(chunk: &mut Chunk, argc: u8, line: u32) {
    if argc < 2 {
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let n = (argc as u16) - 1;
    let arr_slot = chunk.alloc_scratch(2);
    let idx_slot = arr_slot + 1;

    chunk.emit_array_new_fixed(0, n, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    chunk.emit_op(Op::I32_FROM_F64, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

/// `.NET object.GetType()` — return a Type-like descriptor with at least `Name`.
/// Stack: `[value] -> [{ Name = value.name || value.__type || typeof(value) }]`.
fn emit_get_type(chunk: &mut Chunk, line: u32) {
    let value_slot = chunk.alloc_scratch(3);
    let name_slot = value_slot + 1;
    let out_slot = value_slot + 2;
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);

    let name_key = chunk.add_constant(Value::String(Arc::from("name")));
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, name_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);

    emit_slot_is_nullish(chunk, name_slot, line);
    let use_name_block = chunk.emit_block(line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(0, line);
    let type_key = chunk.add_constant(Value::String(Arc::from("__type")));
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, type_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunk.emit_end(line);
    chunk.patch_block(use_name_block);

    emit_slot_is_nullish(chunk, name_slot, line);
    let use_type_block = chunk.emit_block(line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    let typeof_idx = chunk.add_import("ecma:value", "typeof");
    chunk.emit_call(typeof_idx, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunk.emit_end(line);
    chunk.patch_block(use_type_block);

    chunk.emit_op_u16(Op::LOCAL_GET, name_slot, line);
    chunk.emit_string_const("object", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);
    let exception_key = chunk.add_constant(Value::String(Arc::from("__exception_type")));
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, exception_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, name_slot, line);
    chunk.emit_string_const("", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, exception_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunk.emit_end(line);

    chunk.emit_struct_new(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, name_slot, line);
    let public_name_key = chunk.add_constant(Value::String(Arc::from("Name")));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, public_name_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

fn emit_slot_is_nullish(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    let undef = chunk.add_import("wasm:js-undefined", "test");
    chunk.emit_call(undef, 1, line);
    chunk.emit_op(Op::I32_OR, line);
}

fn emit_to_byte(chunk: &mut Chunk, line: u32) {
    let value_slot = chunk.alloc_scratch(1);
    let number_idx = chunk.add_import("ecma:number", "Number");
    chunk.emit_call(number_idx, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_f64_const(255.0, line);
    chunk.emit_op(Op::F64_GT, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if(line);
    chunk.emit_struct_new(0, 0, line);
    chunk.emit_dup(line);
    chunk.emit_string_const("Arithmetic operation resulted in an overflow.", line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(
        chunk,
        "OverflowException",
        line,
    );
    vybe_compiler::primitives::errors::emit_stamp_exception_ancestors(
        chunk,
        "OverflowException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op(Op::I32_FROM_F64, line);
}

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    if crate::emitter::core::runtime_adapter::emit_helper(name, chunks, current, argc, line) {
        return true;
    }
    match name {
        // A member that IS the receiver — `MenuStrip.Items`. WinForms wraps a
        // strip's contents in a collection object, but in the document the
        // element already IS that container, so the getter yields what it was
        // handed and allocates nothing. Emitting NOTHING is the whole
        // implementation: the receiver is already on the stack. plib spells the
        // same member `pascal.self` for the same reason.
        "dotnet.self" => {}
        "dotnet.winforms_application_run" => {
            crate::emitter::winforms::adapter::emit_application_run(chunks, current, argc, line);
        }
        "dotnet.winforms_application_exit" => {
            crate::emitter::winforms::adapter::emit_application_exit(chunks, current, argc, line);
        }
        "dotnet.winforms_noop" => {
            crate::emitter::winforms::adapter::emit_noop(chunks, current, argc, line);
        }
        "dotnet.winforms_message_box_show" => {
            crate::emitter::winforms::adapter::emit_message_box_show(chunks, current, argc, line);
        }
        "dotnet.get_type" => emit_get_type(&mut chunks[current], line),
        "dotnet.to_byte" => emit_to_byte(&mut chunks[current], line),
        "dotnet.winforms_control_show" => {
            crate::emitter::winforms::adapter::emit_control_show(chunks, current, argc, line);
        }
        "dotnet.winforms_control_hide" => {
            crate::emitter::winforms::adapter::emit_control_hide(chunks, current, argc, line);
        }
        "dotnet.winforms_control_close" => {
            crate::emitter::winforms::adapter::emit_control_close(chunks, current, argc, line);
        }
        "dotnet.winforms_form_show_dialog" => {
            crate::emitter::winforms::adapter::emit_form_show_dialog(chunks, current, argc, line);
        }
        "dotnet.winforms_controls_add" => {
            crate::emitter::winforms::adapter::emit_controls_add(chunks, current, argc, line);
        }
        "dotnet.control_create_graphics" => {
            emit_control_create_graphics(&mut chunks[current], line);
        }
        // Drawing methods (Graphics/Pen/Brush) lower their `Body` op sequence
        // inline at the call site, reusing the same MethodOp table that builds
        // a thunk chunk — no per-class ctor chunk binds a thunk anymore.
        drawing if drawing.starts_with("dotnet.drawing.") => {
            let method = &drawing["dotnet.drawing.".len()..];
            match crate::emitter::classes::drawing::drawing_method_body(method) {
                Some(ops) => crate::emitter::classes::builder::emit_body_inline(
                    &mut chunks[current],
                    ops,
                    argc,
                    line,
                ),
                None => return false,
            }
        }
        "dotnet.dns_get_host_addresses" => {
            crate::emitter::core::sockets_adapter::emit_dns_get_host_addresses(
                chunks, current, line,
            )
        }
        "dotnet.dns_get_host_entry" => {
            crate::emitter::core::sockets_adapter::emit_dns_get_host_entry(chunks, current, line)
        }
        "dotnet.dns_get_host_name" => {
            crate::emitter::core::sockets_adapter::emit_dns_get_host_name(chunks, current, line)
        }
        "dotnet.bitconverter_get_bytes" => {
            crate::emitter::core::bitconverter_adapter::emit_get_bytes(chunks, current, line)
        }
        "dotnet.bitconverter_to_number" => {
            crate::emitter::core::bitconverter_adapter::emit_to_number(chunks, current, line)
        }
        "dotnet.bitconverter_to_boolean" => {
            crate::emitter::core::bitconverter_adapter::emit_to_boolean(chunks, current, line)
        }
        "dotnet.bitconverter_to_char" => {
            crate::emitter::core::bitconverter_adapter::emit_to_char(chunks, current, line)
        }
        "dotnet.bitconverter_to_string" => {
            crate::emitter::core::bitconverter_adapter::emit_to_string(chunks, current, line)
        }
        "dotnet.bitconverter_is_little_endian" => {
            crate::emitter::core::bitconverter_adapter::emit_is_little_endian(chunks, current, line)
        }
        "dotnet.buffer_block_copy" => {
            crate::emitter::core::bitconverter_adapter::emit_block_copy(chunks, current, line)
        }
        "dotnet.ip_address_parse" => {
            crate::emitter::core::sockets_adapter::emit_ip_address_parse(chunks, current, line)
        }
        "dotnet.ip_address_to_string" => {
            crate::emitter::core::sockets_adapter::emit_ip_address_to_string(chunks, current, line)
        }
        "dotnet.tcp_client_new" => {
            crate::emitter::core::sockets_adapter::emit_tcp_client_new(chunks, current, line)
        }
        "dotnet.tcp_client_get_stream" => {
            crate::emitter::core::sockets_adapter::emit_tcp_client_get_stream(chunks, current, line)
        }
        "dotnet.tcp_client_close" => {
            crate::emitter::core::sockets_adapter::emit_tcp_client_close(chunks, current, line)
        }
        "dotnet.tcp_listener_new" => {
            crate::emitter::core::sockets_adapter::emit_tcp_listener_new(chunks, current, line)
        }
        "dotnet.tcp_listener_start" => {
            crate::emitter::core::sockets_adapter::emit_tcp_listener_start(chunks, current, line)
        }
        "dotnet.tcp_listener_stop" => {
            crate::emitter::core::sockets_adapter::emit_tcp_listener_stop(chunks, current, line)
        }
        "dotnet.tcp_listener_accept" => {
            crate::emitter::core::sockets_adapter::emit_tcp_listener_accept(chunks, current, line)
        }
        "dotnet.tcp_listener_pending" => {
            crate::emitter::core::sockets_adapter::emit_tcp_listener_pending(chunks, current, line)
        }
        "dotnet.udp_client_new" => {
            crate::emitter::core::sockets_adapter::emit_udp_client_new(chunks, current, line)
        }
        "dotnet.udp_send" => {
            crate::emitter::core::sockets_adapter::emit_udp_send(chunks, current, line)
        }
        "dotnet.udp_receive" => {
            crate::emitter::core::sockets_adapter::emit_udp_receive(chunks, current, line)
        }
        "dotnet.udp_close" => {
            crate::emitter::core::sockets_adapter::emit_udp_close(chunks, current, line)
        }

        // ── .NET StringBuilder adapter ──────────────────────────────
        // No direct ECMA mirror; the wrapper materializes a plain
        // Object with a `__buffer` string and mutates via DYN_ADD +
        // STRUCT_SET. Multi-arity ctor uses the threaded `argc` to
        // pick between empty / initial-keyed shapes.
        "dotnet.string_builder_new" => {
            crate::emitter::core::stringbuilder_adapter::emit_string_builder_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.sb_append" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_append(chunks, current, argc, line)
        }
        "dotnet.sb_append_line" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_append_line(
                chunks, current, argc, line,
            )
        }
        "dotnet.sb_append_format" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_append_format(
                chunks, current, argc, line,
            )
        }
        "dotnet.sb_append_join" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_append_join(chunks, current, line)
        }
        "dotnet.sb_to_string" => crate::emitter::core::stringbuilder_adapter::emit_sb_to_string(
            chunks, current, argc, line,
        ),
        "dotnet.sb_clear" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_clear(chunks, current, line)
        }
        "dotnet.sb_length" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_length(chunks, current, line)
        }
        "dotnet.sb_capacity" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_capacity(chunks, current, line)
        }
        "dotnet.sb_set_capacity" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_set_capacity(chunks, current, line)
        }
        "dotnet.sb_max_capacity" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_max_capacity(chunks, current, line)
        }
        "dotnet.sb_set_length" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_set_length(chunks, current, line)
        }
        "dotnet.sb_ensure_capacity" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_ensure_capacity(
                chunks, current, line,
            )
        }
        "dotnet.sb_copy_to" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_copy_to(chunks, current, line)
        }
        "dotnet.sb_equals" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_equals(chunks, current, line)
        }
        "dotnet.sb_insert" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_insert(chunks, current, line)
        }
        "dotnet.sb_remove" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_remove(chunks, current, line)
        }
        "dotnet.sb_replace" => crate::emitter::core::stringbuilder_adapter::emit_sb_replace(
            chunks, current, argc, line,
        ),
        "dotnet.sb_index_get" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_index_get(chunks, current, line)
        }
        "dotnet.sb_index_set" => {
            crate::emitter::core::stringbuilder_adapter::emit_sb_index_set(chunks, current, line)
        }

        // ── .NET Random adapter ─────────────────────────────────────
        "dotnet.random_new" => {
            crate::emitter::core::random_adapter::emit_random_new(chunks, current, argc, line)
        }
        "dotnet.random_next" => {
            crate::emitter::core::random_adapter::emit_random_next(chunks, current, argc, line)
        }
        "dotnet.random_next_double" => {
            crate::emitter::core::random_adapter::emit_random_next_double(chunks, current, line)
        }
        "dotnet.random_next_bytes" => {
            crate::emitter::core::random_adapter::emit_random_next_bytes(chunks, current, line)
        }

        // ── .NET Regex adapter ──────────────────────────────────────
        "dotnet.regex_new" => {
            crate::emitter::core::regex_adapter::emit_regex_new(chunks, current, argc, line)
        }
        "dotnet.regex_is_match" => {
            crate::emitter::core::regex_adapter::emit_regex_is_match(chunks, current, line)
        }
        "dotnet.regex_static_is_match" => {
            crate::emitter::core::regex_adapter::emit_regex_static_is_match(
                chunks, current, argc, line,
            )
        }
        "dotnet.regex_static_match" => {
            crate::emitter::core::regex_adapter::emit_regex_static_match(
                chunks, current, argc, line,
            )
        }
        "dotnet.regex_static_matches" => {
            crate::emitter::core::regex_adapter::emit_regex_static_matches(
                chunks, current, argc, line,
            )
        }
        "dotnet.regex_static_replace" => {
            crate::emitter::core::regex_adapter::emit_regex_static_replace(
                chunks, current, argc, line,
            )
        }
        "dotnet.regex_static_split" => {
            crate::emitter::core::regex_adapter::emit_regex_static_split(
                chunks, current, argc, line,
            )
        }
        "dotnet.regex_escape" => {
            crate::emitter::core::regex_adapter::emit_regex_escape(chunks, current, line)
        }
        "dotnet.regex_unescape" => {
            crate::emitter::core::regex_adapter::emit_regex_unescape(chunks, current, line)
        }
        "dotnet.regex_replace" => {
            crate::emitter::core::regex_adapter::emit_regex_replace(chunks, current, line)
        }
        "dotnet.regex_split" => {
            crate::emitter::core::regex_adapter::emit_regex_split(chunks, current, line)
        }
        "dotnet.regex_match" => {
            crate::emitter::core::regex_adapter::emit_regex_match(chunks, current, line)
        }
        "dotnet.regex_matches" => {
            crate::emitter::core::regex_adapter::emit_regex_matches(chunks, current, line)
        }
        "dotnet.regex_get_group_names" => {
            crate::emitter::core::regex_adapter::emit_regex_get_group_names(chunks, current, line)
        }
        "dotnet.regex_group_name_from_number" => {
            crate::emitter::core::regex_adapter::emit_regex_group_name_from_number(
                chunks, current, line,
            )
        }
        "dotnet.regex_group_number_from_name" => {
            crate::emitter::core::regex_adapter::emit_regex_group_number_from_name(
                chunks, current, line,
            )
        }

        // ── .NET Stopwatch adapter ──────────────────────────────────
        // System.Net — lowered onto the real WASI 0.2 HTTP interfaces
        // (request -> outgoing-handler.handle -> consume-body). Replaces the
        // old `wasi:http.fetch` target, which named an unregistered module.
        "dotnet.http_fetch" => {
            crate::emitter::core::http_adapter::emit_http_fetch(chunks, current, line)
        }
        "dotnet.http_client_new" => {
            crate::emitter::core::http_adapter::emit_http_client_new(chunks, current, line)
        }
        "dotnet.stopwatch_new" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_new(chunks, current, line)
        }
        "dotnet.stopwatch_start" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_start(chunks, current, line)
        }
        "dotnet.stopwatch_stop" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_stop(chunks, current, line)
        }
        "dotnet.stopwatch_reset" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_reset(chunks, current, line)
        }
        "dotnet.stopwatch_start_new" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_start_new(chunks, current, line)
        }
        "dotnet.stopwatch_restart" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_restart(chunks, current, line)
        }
        "dotnet.stopwatch_elapsed_ms" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_elapsed_ms(
                chunks, current, line,
            )
        }
        "dotnet.stopwatch_elapsed_ticks" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_elapsed_ticks(
                chunks, current, line,
            )
        }
        "dotnet.stopwatch_elapsed" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_elapsed(chunks, current, line)
        }
        "dotnet.stopwatch_is_running" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_is_running(
                chunks, current, line,
            )
        }
        "dotnet.stopwatch_frequency" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_frequency(chunks, current, line)
        }
        "dotnet.stopwatch_is_high_resolution" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_is_high_resolution(
                chunks, current, line,
            )
        }
        "dotnet.stopwatch_get_timestamp" => {
            crate::emitter::core::stopwatch_adapter::emit_stopwatch_get_timestamp(
                chunks, current, line,
            )
        }

        // ── .NET Process / ProcessStartInfo adapter ─────────────────
        // Lowers to `node:child_process.spawnSync` + plain Object
        // structs for the .NET-shape records. Multi-arity ctors use
        // the threaded `argc`.
        "dotnet.process_start_info_new" => {
            crate::emitter::core::process_adapter::emit_process_start_info_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.process_new" => {
            crate::emitter::core::process_adapter::emit_process_new(chunks, current, argc, line)
        }
        "dotnet.process_start" => {
            crate::emitter::core::process_adapter::emit_process_start(chunks, current, line)
        }
        "dotnet.process_get_current" => {
            crate::emitter::core::process_adapter::emit_process_get_current(chunks, current, line)
        }
        "dotnet.process_get_by_id" => {
            crate::emitter::core::process_adapter::emit_process_get_by_id(chunks, current, line)
        }
        "dotnet.process_get_processes" => {
            crate::emitter::core::process_adapter::emit_process_get_processes(chunks, current, line)
        }
        "dotnet.process_get_processes_by_name" => {
            crate::emitter::core::process_adapter::emit_process_get_processes_by_name(
                chunks, current, line,
            )
        }
        "dotnet.process_wait_for_exit" => {
            crate::emitter::core::process_adapter::emit_process_wait_for_exit(chunks, current, line)
        }
        "dotnet.process_wait_for_exit_timeout" => {
            crate::emitter::core::process_adapter::emit_process_wait_for_exit_timeout(
                chunks, current, line,
            )
        }
        "dotnet.enum_is_defined" => {
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_op(Op::DROP, line);
            vybe_compiler::primitives::instructions::core_wasm::bool_const(
                &mut chunks[current],
                line,
                true,
            );
        }

        // ── .NET LINQ-to-XML adapter ────────────────────────────────
        // Tree operations compose `web:dom-parser`; XName operations delegate
        // to `vybe_compiler::primitives::xml` inside the adapter.
        "dotnet.xml_xelement_new" => {
            crate::emitter::core::xml_linq_adapter::emit_xelement_new(chunks, current, argc, line)
        }
        "dotnet.xml_xattribute_new" => {
            crate::emitter::core::xml_linq_adapter::emit_xattribute_new(chunks, current, argc, line)
        }
        "dotnet.xml_xdocument_new" => {
            crate::emitter::core::xml_linq_adapter::emit_xdocument_new(chunks, current, argc, line)
        }
        "dotnet.xml_xdocument_parse" => {
            crate::emitter::core::xml_linq_adapter::emit_xdocument_parse(chunks, current, line)
        }
        "dotnet.xml_xdocument_root" => {
            crate::emitter::core::xml_linq_adapter::emit_xdocument_root(chunks, current, line)
        }
        "dotnet.xml_xelement_name" => {
            crate::emitter::core::xml_linq_adapter::emit_xelement_name(chunks, current, line)
        }
        "dotnet.xml_value" => {
            crate::emitter::core::xml_linq_adapter::emit_xml_value(chunks, current, line)
        }
        "dotnet.xml_to_string" => {
            crate::emitter::core::xml_linq_adapter::emit_xml_to_string(chunks, current, line)
        }
        "dotnet.xml_element" => {
            crate::emitter::core::xml_linq_adapter::emit_xml_element(chunks, current, line)
        }
        "dotnet.xml_child_elements" => {
            crate::emitter::core::xml_linq_adapter::emit_xml_child_elements(chunks, current, line)
        }
        "dotnet.xml_elements" => {
            crate::emitter::core::xml_linq_adapter::emit_xml_elements(chunks, current, line)
        }
        "dotnet.xml_attribute" => {
            crate::emitter::core::xml_linq_adapter::emit_xml_attribute(chunks, current, line)
        }
        "dotnet.xml_attribute_value" => {
            crate::emitter::core::xml_linq_adapter::emit_attribute_value(chunks, current, line)
        }
        "dotnet.xml_xelement_add" => {
            crate::emitter::core::xml_linq_adapter::emit_xelement_add(chunks, current, line)
        }
        "dotnet.xml_xelement_remove" => {
            crate::emitter::core::xml_linq_adapter::emit_xelement_remove(chunks, current, line)
        }
        "dotnet.xml_xelement_replace_nodes" => {
            crate::emitter::core::xml_linq_adapter::emit_xelement_replace_nodes(
                chunks, current, line,
            )
        }
        "dotnet.xml_xelement_set_attribute_value" => {
            crate::emitter::core::xml_linq_adapter::emit_xelement_set_attribute_value(
                chunks, current, line,
            )
        }

        // ── .NET System.Array static-method adapter ─────────────────
        // `Clear` / `Copy` / `Resize` / `Sort` lower to bundled stdlib
        // chunks (`__vybe_*` globals) composing `ecma:array.*`
        // primitives. No `vybe:types/array*` host fns.
        "dotnet.array_clear" => {
            crate::emitter::core::array_adapter::emit_array_clear(chunks, current, line)
        }
        "dotnet.array_copy" => {
            crate::emitter::core::array_adapter::emit_array_copy(chunks, current, argc, line)
        }
        "dotnet.array_get_checked" => {
            crate::emitter::core::array_adapter::emit_array_get_checked(chunks, current, line)
        }
        "dotnet.array_set_checked" => {
            crate::emitter::core::array_adapter::emit_array_set_checked(chunks, current, line)
        }
        "dotnet.list_get_checked" => {
            crate::emitter::core::array_adapter::emit_list_get_checked(chunks, current, line)
        }
        "dotnet.get_range_checked" => {
            crate::emitter::core::array_adapter::emit_get_range_checked(chunks, current, line)
        }
        "dotnet.array_resize" => {
            crate::emitter::core::array_adapter::emit_array_resize(chunks, current, line)
        }
        "dotnet.array_sort" => {
            crate::emitter::core::array_adapter::emit_array_sort(chunks, current, line)
        }
        "dotnet.hashset_add" => {
            crate::emitter::core::collections_adapter::emit_hashset_add(chunks, current, line)
        }
        "dotnet.list_new" => {
            crate::emitter::core::collections_adapter::emit_list_new(chunks, current, argc, line)
        }
        "dotnet.list_add" => {
            crate::emitter::core::collections_adapter::emit_list_add(chunks, current, line)
        }
        "dotnet.list_remove_all" => {
            crate::emitter::core::array_adapter::emit_list_remove_all(chunks, current, line)
        }
        "dotnet.list_capacity" => {
            crate::emitter::core::collections_adapter::emit_list_capacity(chunks, current, line)
        }
        "dotnet.list_ensure_capacity" => {
            crate::emitter::core::collections_adapter::emit_list_ensure_capacity(
                chunks, current, line,
            )
        }
        "dotnet.list_trim_excess" => {
            crate::emitter::core::collections_adapter::emit_list_trim_excess(chunks, current, line)
        }
        "dotnet.set_new_ignore_comparer" => {
            crate::emitter::core::collections_adapter::emit_set_new_ignore_comparer(
                chunks, current, line,
            )
        }
        "dotnet.list_new_from_iterable" => {
            crate::emitter::core::collections_adapter::emit_list_new_from_iterable(
                chunks, current, line,
            )
        }
        "dotnet.readonly_observable_collection_new" => {
            crate::emitter::core::collections_adapter::emit_readonly_observable_collection_new(
                chunks, current, line,
            )
        }
        "dotnet.property_changed_event_args_new" => {
            crate::emitter::core::collections_adapter::emit_property_changed_event_args_new(
                chunks, current, line,
            )
        }
        "dotnet.notify_collection_changed_event_args_new" => {
            crate::emitter::core::collections_adapter::emit_notify_collection_changed_event_args_new(
                chunks, current, line,
            )
        }
        "dotnet.set_new_from_iterable" => {
            crate::emitter::core::collections_adapter::emit_set_new_from_iterable(
                chunks, current, line,
            )
        }
        "dotnet.hashset_union_with" => {
            crate::emitter::core::collections_adapter::emit_hashset_union_with(
                chunks, current, line,
            )
        }
        "dotnet.hashset_intersect_with" => {
            crate::emitter::core::collections_adapter::emit_hashset_intersect_with(
                chunks, current, line,
            )
        }
        "dotnet.hashset_except_with" => {
            crate::emitter::core::collections_adapter::emit_hashset_except_with(
                chunks, current, line,
            )
        }
        "dotnet.hashset_symmetric_except_with" => {
            crate::emitter::core::collections_adapter::emit_hashset_symmetric_except_with(
                chunks, current, line,
            )
        }
        "dotnet.hashset_is_subset_of" => {
            crate::emitter::core::collections_adapter::emit_hashset_is_subset_of(
                chunks, current, line,
            )
        }
        "dotnet.hashset_is_superset_of" => {
            crate::emitter::core::collections_adapter::emit_hashset_is_superset_of(
                chunks, current, line,
            )
        }
        "dotnet.hashset_overlaps" => {
            crate::emitter::core::collections_adapter::emit_hashset_overlaps(chunks, current, line)
        }
        "dotnet.task_wait" => {
            crate::emitter::core::thread_adapter::emit_task_wait(chunks, current, line)
        }
        "dotnet.task_run" => {
            crate::emitter::core::thread_adapter::emit_task_run(chunks, current, line)
        }
        "dotnet.task_delay" => {
            crate::emitter::core::thread_adapter::emit_task_delay(chunks, current, argc, line)
        }
        "dotnet.task_from_result" => {
            crate::emitter::core::thread_adapter::emit_task_from_result(chunks, current, line)
        }
        "dotnet.task_when_all" => {
            crate::emitter::core::thread_adapter::emit_task_when_all(chunks, current, argc, line)
        }
        "dotnet.task_when_any" => {
            crate::emitter::core::thread_adapter::emit_task_when_any(chunks, current, argc, line)
        }
        "dotnet.task_continue_with" => {
            crate::emitter::core::thread_adapter::emit_task_continue_with(chunks, current, line)
        }
        "dotnet.task_yield" => {
            crate::emitter::core::thread_adapter::emit_task_yield(chunks, current, line)
        }
        "dotnet.value_task_as_task" => {
            crate::emitter::core::thread_adapter::emit_value_task_as_task(chunks, current, line)
        }
        "dotnet.task_is_canceled" => {
            crate::emitter::core::thread_adapter::emit_task_is_canceled(chunks, current, line)
        }
        "dotnet.task_result" => {
            crate::emitter::core::thread_adapter::emit_task_result(chunks, current, line)
        }
        "dotnet.task_is_completed" => {
            crate::emitter::core::thread_adapter::emit_task_is_completed(chunks, current, line)
        }
        "dotnet.noop" => {
            for _ in 0..argc {
                chunks[current].emit_op(Op::DROP, line);
            }
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        }
        "dotnet.cancellation_token_source_new" => {
            crate::emitter::core::thread_adapter::emit_cancellation_token_source_new(
                chunks, current, line,
            )
        }
        "dotnet.cancellation_token_none" => {
            crate::emitter::core::thread_adapter::emit_cancellation_token_none(
                chunks, current, line,
            )
        }
        "dotnet.cancellation_token_cancel" => {
            crate::emitter::core::thread_adapter::emit_cancellation_token_cancel(
                chunks, current, line,
            )
        }
        "dotnet.cancellation_token_cancel_after" => {
            crate::emitter::core::thread_adapter::emit_cancellation_token_cancel_after(
                chunks, current, line,
            )
        }
        "dotnet.cancellation_token_source_token" => {
            crate::emitter::core::thread_adapter::emit_cancellation_token_source_token(
                chunks, current, line,
            )
        }
        "dotnet.cancellation_token_is_requested" => {
            crate::emitter::core::thread_adapter::emit_cancellation_token_is_requested(
                chunks, current, line,
            )
        }
        "dotnet.cancellation_token_throw_if_requested" => {
            crate::emitter::core::thread_adapter::emit_cancellation_token_throw_if_requested(
                chunks, current, line,
            )
        }
        "dotnet.cancellation_token_register" => {
            crate::emitter::core::thread_adapter::emit_cancellation_token_register(
                chunks, current, line,
            )
        }
        "dotnet.cancellation_token_can_be_canceled" => {
            crate::emitter::core::thread_adapter::emit_cancellation_token_can_be_canceled(
                chunks, current, line,
            )
        }
        "dotnet.cancellation_token_wait_handle" => {
            crate::emitter::core::thread_adapter::emit_cancellation_token_wait_handle(
                chunks, current, line,
            )
        }
        "dotnet.cancellation_token_linked_source" => {
            crate::emitter::core::thread_adapter::emit_cancellation_token_linked_source(
                chunks, current, argc, line,
            )
        }
        "dotnet.hashset_set_equals" => {
            crate::emitter::core::collections_adapter::emit_hashset_set_equals(
                chunks, current, line,
            )
        }
        "dotnet.hashset_is_proper_subset_of" => {
            crate::emitter::core::collections_adapter::emit_hashset_is_proper_subset_of(
                chunks, current, line,
            )
        }
        "dotnet.hashset_is_proper_superset_of" => {
            crate::emitter::core::collections_adapter::emit_hashset_is_proper_superset_of(
                chunks, current, line,
            )
        }
        "dotnet.linked_list_add_first" => {
            crate::emitter::core::collections_adapter::emit_linked_list_add_first(
                chunks, current, line,
            )
        }
        "dotnet.linked_list_add_last" => {
            crate::emitter::core::collections_adapter::emit_linked_list_add_last(
                chunks, current, line,
            )
        }
        "dotnet.linked_list_find" => {
            crate::emitter::core::collections_adapter::emit_linked_list_find(chunks, current, line)
        }
        "dotnet.linked_list_first" => {
            crate::emitter::core::collections_adapter::emit_linked_list_first(chunks, current, line)
        }
        "dotnet.vb_collection_new" => {
            crate::emitter::core::collections_adapter::emit_vb_collection_new(chunks, current, line)
        }
        "dotnet.vb_collection_add" => {
            crate::emitter::core::collections_adapter::emit_vb_collection_add(
                chunks, current, argc, line,
            )
        }
        "dotnet.vb_collection_item" => {
            crate::emitter::core::collections_adapter::emit_vb_collection_item(
                chunks, current, line,
            )
        }
        "dotnet.vb_collection_count" => {
            crate::emitter::core::collections_adapter::emit_vb_collection_count(
                chunks, current, line,
            )
        }
        "dotnet.vb_collection_to_array" => {
            crate::emitter::core::collections_adapter::emit_vb_collection_to_array(
                chunks, current, line,
            )
        }
        "dotnet.vb_collection_contains" => {
            crate::emitter::core::collections_adapter::emit_vb_collection_contains(
                chunks, current, line,
            )
        }
        "dotnet.vb_collection_remove" => {
            crate::emitter::core::collections_adapter::emit_vb_collection_remove(
                chunks, current, line,
            )
        }
        "dotnet.dict_new_ignore_arg" => {
            crate::emitter::core::collections_adapter::emit_dict_new_ignore_arg(
                chunks, current, line,
            )
        }
        "dotnet.dict_try_add" => {
            crate::emitter::core::collections_adapter::emit_dict_try_add(chunks, current, line)
        }
        "dotnet.dict_remove" => {
            crate::emitter::core::collections_adapter::emit_dict_remove(chunks, current, line)
        }
        "dotnet.dict_ensure_capacity" => {
            crate::emitter::core::collections_adapter::emit_dict_ensure_capacity(
                chunks, current, line,
            )
        }
        "dotnet.dict_trim_excess" => {
            crate::emitter::core::collections_adapter::emit_dict_trim_excess(chunks, current, line)
        }
        "dotnet.blocking_collection_new" => {
            crate::emitter::core::collections_adapter::emit_blocking_collection_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.blocking_collection_add" => {
            crate::emitter::core::collections_adapter::emit_blocking_collection_add(
                chunks, current, line,
            )
        }
        "dotnet.blocking_collection_try_add" => {
            crate::emitter::core::collections_adapter::emit_blocking_collection_try_add(
                chunks, current, line,
            )
        }
        "dotnet.blocking_collection_take" => {
            crate::emitter::core::collections_adapter::emit_blocking_collection_take(
                chunks, current, argc, line,
            )
        }
        "dotnet.blocking_collection_count" => {
            crate::emitter::core::collections_adapter::emit_blocking_collection_count(
                chunks, current, line,
            )
        }
        "dotnet.blocking_collection_complete_adding" => {
            crate::emitter::core::collections_adapter::emit_blocking_collection_complete_adding(
                chunks, current, line,
            )
        }
        "dotnet.blocking_collection_is_completed" => {
            crate::emitter::core::collections_adapter::emit_blocking_collection_is_completed(
                chunks, current, line,
            )
        }
        "dotnet.blocking_collection_items" => {
            crate::emitter::core::collections_adapter::emit_blocking_collection_items(
                chunks, current, line,
            )
        }
        "dotnet.observable_collection_items" => {
            crate::emitter::core::collections_adapter::emit_observable_collection_items(
                chunks, current, line,
            )
        }
        "dotnet.observable_collection_count" => {
            crate::emitter::core::collections_adapter::emit_observable_collection_count(
                chunks, current, line,
            )
        }
        "dotnet.observable_collection_add" => {
            crate::emitter::core::collections_adapter::emit_observable_collection_add(
                chunks, current, line,
            )
        }
        "dotnet.observable_collection_remove" => {
            crate::emitter::core::collections_adapter::emit_observable_collection_remove(
                chunks, current, line,
            )
        }
        "dotnet.observable_collection_remove_at" => {
            crate::emitter::core::collections_adapter::emit_observable_collection_remove_at(
                chunks, current, line,
            )
        }
        "dotnet.observable_collection_insert" => {
            crate::emitter::core::collections_adapter::emit_observable_collection_insert(
                chunks, current, line,
            )
        }
        "dotnet.observable_collection_set_index" => {
            crate::emitter::core::collections_adapter::emit_observable_collection_set_index(
                chunks, current, line,
            )
        }
        "dotnet.observable_collection_move" => {
            crate::emitter::core::collections_adapter::emit_observable_collection_move(
                chunks, current, line,
            )
        }
        "dotnet.observable_collection_clear" => {
            crate::emitter::core::collections_adapter::emit_observable_collection_clear(
                chunks, current, line,
            )
        }
        "dotnet.observable_collection_on_changed" => {
            crate::emitter::core::collections_adapter::emit_observable_collection_on_changed(
                chunks, current, line,
            )
        }
        "dotnet.sorted_dictionary_entries" => {
            crate::emitter::core::collections_adapter::emit_sorted_dictionary_entries(
                chunks, current, line,
            )
        }

        // ── Sorted collections (SortedSet / SortedDictionary) ────────
        // Both keep their `ecma:set` / `ecma:map` backing (so membership,
        // mutation and set-algebra reuse the host ops); only the ordered reads
        // are adapted through the shared sorted core in
        // `vybe_compiler::primitives::sorted_collection`, which is also Java's TreeSet/TreeMap
        // engine. A SortedSet spreads to a sorted array for Min/Max/GetViewBetween
        // and ElementsSorted; a SortedDictionary sorts its key/value/entry views.
        "dotnet.sorted_set_min" => {
            crate::emitter::core::collections_adapter::emit_sorted_set_min(chunks, current, line)
        }
        "dotnet.sorted_set_max" => {
            crate::emitter::core::collections_adapter::emit_sorted_set_max(chunks, current, line)
        }
        "dotnet.sorted_set_elements" => {
            crate::emitter::core::collections_adapter::emit_sorted_set_elements(
                chunks, current, line,
            )
        }
        "dotnet.sorted_set_view_between" => {
            crate::emitter::core::collections_adapter::emit_sorted_set_view_between(
                chunks, current, line,
            )
        }
        "dotnet.sorted_map_keys" => {
            vybe_compiler::primitives::sorted_collection::emit_sorted_map_key_set(
                chunks, current, line,
            )
        }
        "dotnet.sorted_map_values" => {
            vybe_compiler::primitives::sorted_collection::emit_sorted_map_values(
                chunks, current, line,
            )
        }
        "dotnet.sorted_map_entries" => {
            vybe_compiler::primitives::sorted_collection::emit_sorted_map_entries(
                chunks, current, line,
            )
        }

        // ── .NET TimeSpan factory adapters ──────────────────────────
        // `TimeSpan.From*(n)` factories build a duration record by
        // multiplying `n` with the unit-to-ms factor. Pure inline
        // bytecode; no host fns.
        "dotnet.timespan_from_days" => {
            crate::emitter::core::timespan_adapter::emit_timespan_from_days(chunks, current, line)
        }
        "dotnet.timespan_from_hours" => {
            crate::emitter::core::timespan_adapter::emit_timespan_from_hours(chunks, current, line)
        }
        "dotnet.timespan_from_minutes" => {
            crate::emitter::core::timespan_adapter::emit_timespan_from_minutes(
                chunks, current, line,
            )
        }
        "dotnet.timespan_from_seconds" => {
            crate::emitter::core::timespan_adapter::emit_timespan_from_seconds(
                chunks, current, line,
            )
        }
        "dotnet.timespan_from_milliseconds" => {
            crate::emitter::core::timespan_adapter::emit_timespan_from_milliseconds(
                chunks, current, line,
            )
        }
        "dotnet.timespan_zero" => {
            crate::emitter::core::timespan_adapter::emit_timespan_zero(chunks, current, line)
        }
        "dotnet.timespan_new" => {
            crate::emitter::core::timespan_adapter::emit_timespan_new(chunks, current, argc, line)
        }
        "dotnet.timespan_compare" => {
            crate::emitter::core::timespan_adapter::emit_timespan_compare(chunks, current, line)
        }
        "dotnet.timespan_parse" => {
            crate::emitter::core::timespan_adapter::emit_timespan_parse(chunks, current, line)
        }
        "dotnet.timespan_negate" => {
            crate::emitter::core::timespan_adapter::emit_timespan_negate(chunks, current, line)
        }
        "dotnet.timespan_duration" => {
            crate::emitter::core::timespan_adapter::emit_timespan_duration(chunks, current, line)
        }
        "dotnet.timespan_add" => {
            crate::emitter::core::timespan_adapter::emit_timespan_add(chunks, current, line)
        }
        "dotnet.timespan_sub" => {
            crate::emitter::core::timespan_adapter::emit_timespan_sub(chunks, current, line)
        }

        // ── .NET Guid adapters ──────────────────────────────────────
        // `Guid` is stored as a .NET-shaped object carrying the
        // canonical lowercase text representation in `__value`.
        "dotnet.guid_empty" => {
            crate::emitter::core::guid_adapter::emit_guid_empty(chunks, current, line)
        }
        "dotnet.guid_new_guid" => {
            crate::emitter::core::guid_adapter::emit_guid_new_guid(chunks, current, line)
        }
        "dotnet.guid_parse" => {
            crate::emitter::core::guid_adapter::emit_guid_parse(chunks, current, line)
        }
        "dotnet.guid_new" => {
            crate::emitter::core::guid_adapter::emit_guid_new(chunks, current, argc, line)
        }
        "dotnet.guid_to_string" => {
            crate::emitter::core::guid_adapter::emit_guid_to_string(chunks, current, argc, line)
        }
        "dotnet.guid_to_byte_array" => {
            crate::emitter::core::guid_adapter::emit_guid_to_byte_array(chunks, current, line)
        }
        "dotnet.guid_get_hash_code" => {
            crate::emitter::core::guid_adapter::emit_guid_get_hash_code(chunks, current, line)
        }
        "dotnet.guid_try_parse" => {
            crate::emitter::core::guid_adapter::emit_guid_try_parse(chunks, current, argc, line)
        }

        "dotnet.version_new" => {
            crate::emitter::core::version_adapter::emit_version_new(chunks, current, argc, line)
        }
        "dotnet.version_parse" => {
            crate::emitter::core::version_adapter::emit_version_parse(chunks, current, line)
        }
        "dotnet.version_try_parse" => {
            crate::emitter::core::version_adapter::emit_version_try_parse(chunks, current, line)
        }
        "dotnet.version_to_string" => {
            crate::emitter::core::version_adapter::emit_version_to_string(
                chunks, current, argc, line,
            )
        }
        "dotnet.version_clone" => {
            crate::emitter::core::version_adapter::emit_version_clone(chunks, current, line)
        }
        "dotnet.version_compare" => {
            crate::emitter::core::version_adapter::emit_version_compare(chunks, current, line)
        }
        "dotnet.version_compare_instance" => {
            crate::emitter::core::version_adapter::emit_version_compare_instance(
                chunks, current, line,
            )
        }
        "dotnet.version_equals" => {
            crate::emitter::core::version_adapter::emit_version_equals(chunks, current, line)
        }
        "dotnet.version_lt" => {
            crate::emitter::core::version_adapter::emit_version_lt(chunks, current, line)
        }
        "dotnet.version_gt" => {
            crate::emitter::core::version_adapter::emit_version_gt(chunks, current, line)
        }
        "dotnet.version_eq" => {
            crate::emitter::core::version_adapter::emit_version_eq(chunks, current, line)
        }
        "dotnet.version_ne" => {
            crate::emitter::core::version_adapter::emit_version_ne(chunks, current, line)
        }
        "dotnet.uri_new" => {
            crate::emitter::core::uri_adapter::emit_uri_new(chunks, current, argc, line)
        }
        "dotnet.uri_to_string" => {
            crate::emitter::core::uri_adapter::emit_uri_to_string(chunks, current, line)
        }
        "dotnet.uri_escape_data_string" => {
            crate::emitter::core::uri_adapter::emit_uri_escape(chunks, current, line)
        }
        "dotnet.uri_unescape_data_string" => {
            crate::emitter::core::uri_adapter::emit_uri_unescape(chunks, current, line)
        }
        "dotnet.uri_is_well_formed" => {
            crate::emitter::core::uri_adapter::emit_uri_is_well_formed(chunks, current, line)
        }
        "dotnet.uri_try_create" => {
            crate::emitter::core::uri_adapter::emit_uri_try_create(chunks, current, argc, line)
        }
        "dotnet.uri_is_base_of" => {
            crate::emitter::core::uri_adapter::emit_uri_is_base_of(chunks, current, line)
        }
        "dotnet.uri_make_relative" => {
            crate::emitter::core::uri_adapter::emit_uri_make_relative(chunks, current, line)
        }
        "dotnet.uri_kind_relative_or_absolute" => crate::emitter::core::uri_adapter::emit_uri_kind(
            "RelativeOrAbsolute",
            chunks,
            current,
            line,
        ),
        "dotnet.uri_kind_absolute" => {
            crate::emitter::core::uri_adapter::emit_uri_kind("Absolute", chunks, current, line)
        }
        "dotnet.uri_kind_relative" => {
            crate::emitter::core::uri_adapter::emit_uri_kind("Relative", chunks, current, line)
        }

        // ── .NET DateTime static adapters ───────────────────────────
        // `Now` / `UtcNow` / `Today` lower to `ecma:date.now` (which
        // reads `wasi:clocks/wall-clock.now`); `Parse` lowers to
        // `ecma:date.parse`. Each wraps the resulting ms timestamp
        // in a `{__type:"DateTime", __time:ms}` object so the .NET
        // surface looks .NET-shaped.
        "dotnet.datetime_now" => {
            crate::emitter::core::datetime_adapter::emit_datetime_now(chunks, current, line)
        }
        "dotnet.datetime_parse" => {
            crate::emitter::core::datetime_adapter::emit_datetime_parse(chunks, current, line)
        }
        "dotnet.datetime_try_parse" => {
            crate::emitter::core::datetime_adapter::emit_datetime_try_parse(chunks, current, line)
        }
        "dotnet.datetime_parse_exact" => {
            crate::emitter::core::datetime_adapter::emit_datetime_parse_exact(
                chunks, current, argc, line,
            )
        }
        "dotnet.datetime_today" => {
            crate::emitter::core::datetime_adapter::emit_datetime_today(chunks, current, line)
        }
        "dotnet.datetime_min_value" => {
            crate::emitter::core::datetime_adapter::emit_datetime_min_value(chunks, current, line)
        }
        "dotnet.datetime_max_value" => {
            crate::emitter::core::datetime_adapter::emit_datetime_max_value(chunks, current, line)
        }
        "dotnet.datetime_new" => {
            crate::emitter::core::datetime_adapter::emit_datetime_new(chunks, current, argc, line)
        }
        "dotnet.datetime_year" => {
            crate::emitter::core::datetime_adapter::emit_datetime_year(chunks, current, line)
        }
        "dotnet.datetime_month" => {
            crate::emitter::core::datetime_adapter::emit_datetime_month(chunks, current, line)
        }
        "dotnet.datetime_day" => {
            crate::emitter::core::datetime_adapter::emit_datetime_day(chunks, current, line)
        }
        "dotnet.datetime_hour" => {
            crate::emitter::core::datetime_adapter::emit_datetime_hour(chunks, current, line)
        }
        "dotnet.datetime_minute" => {
            crate::emitter::core::datetime_adapter::emit_datetime_minute(chunks, current, line)
        }
        "dotnet.datetime_second" => {
            crate::emitter::core::datetime_adapter::emit_datetime_second(chunks, current, line)
        }
        "dotnet.datetime_millisecond" => {
            crate::emitter::core::datetime_adapter::emit_datetime_millisecond(chunks, current, line)
        }
        "dotnet.datetime_day_of_year" => {
            crate::emitter::core::datetime_adapter::emit_datetime_day_of_year(chunks, current, line)
        }
        "dotnet.datetime_day_of_week" => {
            crate::emitter::core::datetime_adapter::emit_datetime_day_of_week(chunks, current, line)
        }
        "dotnet.datetime_ticks" => {
            crate::emitter::core::datetime_adapter::emit_datetime_ticks(chunks, current, line)
        }
        "dotnet.datetime_kind" => {
            crate::emitter::core::datetime_adapter::emit_datetime_kind(chunks, current, line)
        }
        "dotnet.datetime_date" => {
            crate::emitter::core::datetime_adapter::emit_datetime_date(chunks, current, line)
        }
        "dotnet.datetime_time_of_day" => {
            crate::emitter::core::datetime_adapter::emit_datetime_time_of_day(chunks, current, line)
        }
        "dotnet.datetime_add_days" => {
            crate::emitter::core::datetime_adapter::emit_datetime_add_days(chunks, current, line)
        }
        "dotnet.datetime_add_hours" => {
            crate::emitter::core::datetime_adapter::emit_datetime_add_hours(chunks, current, line)
        }
        "dotnet.datetime_add_months" => {
            crate::emitter::core::datetime_adapter::emit_datetime_add_months(chunks, current, line)
        }
        "dotnet.datetime_add_years" => {
            crate::emitter::core::datetime_adapter::emit_datetime_add_years(chunks, current, line)
        }
        "dotnet.datetime_days_in_month" => {
            crate::emitter::core::datetime_adapter::emit_datetime_days_in_month(
                chunks, current, line,
            )
        }
        "dotnet.datetime_is_leap_year" => {
            crate::emitter::core::datetime_adapter::emit_datetime_is_leap_year(
                chunks, current, line,
            )
        }
        "dotnet.datetime_compare" => {
            crate::emitter::core::datetime_adapter::emit_datetime_compare(chunks, current, line)
        }
        "dotnet.datetime_equals_static" | "dotnet.datetime_equals_instance" => {
            crate::emitter::core::datetime_adapter::emit_datetime_equals(chunks, current, line)
        }
        "dotnet.datetime_to_short_date_string" => {
            crate::emitter::core::datetime_adapter::emit_datetime_to_short_date_string(
                chunks, current, line,
            )
        }
        "dotnet.datetime_to_string" => {
            crate::emitter::core::datetime_adapter::emit_datetime_to_string(
                chunks, current, argc, line,
            )
        }
        "dotnet.datetime_to_universal_time" => {
            crate::emitter::core::datetime_adapter::emit_datetime_to_universal_time(
                chunks, current, line,
            )
        }
        "dotnet.datetime_to_local_time" => {
            crate::emitter::core::datetime_adapter::emit_datetime_to_local_time(
                chunks, current, line,
            )
        }
        "dotnet.datetime_to_binary" => {
            crate::emitter::core::datetime_adapter::emit_datetime_to_binary(chunks, current, line)
        }
        "dotnet.datetime_from_binary" => {
            crate::emitter::core::datetime_adapter::emit_datetime_from_binary(chunks, current, line)
        }
        "dotnet.datetime_to_file_time_utc" => {
            crate::emitter::core::datetime_adapter::emit_datetime_to_file_time_utc(
                chunks, current, line,
            )
        }
        "dotnet.datetime_from_file_time_utc" => {
            crate::emitter::core::datetime_adapter::emit_datetime_from_file_time_utc(
                chunks, current, line,
            )
        }
        "dotnet.datetime_to_oadate" => {
            crate::emitter::core::datetime_adapter::emit_datetime_to_oadate(chunks, current, line)
        }
        "dotnet.datetime_from_oadate" => {
            crate::emitter::core::datetime_adapter::emit_datetime_from_oadate(chunks, current, line)
        }
        "dotnet.datetime_get_hash_code" => {
            crate::emitter::core::datetime_adapter::emit_datetime_get_hash_code(
                chunks, current, line,
            )
        }
        "dotnet.datetime_add_timespan" => {
            crate::emitter::core::datetime_adapter::emit_datetime_add_timespan(
                chunks, current, line,
            )
        }
        "dotnet.datetime_add_ticks" => {
            crate::emitter::core::datetime_adapter::emit_datetime_add_ticks(chunks, current, line)
        }
        "dotnet.datetime_specify_kind" => {
            crate::emitter::core::datetime_adapter::emit_datetime_specify_kind(
                chunks, current, line,
            )
        }
        "dotnet.datetime_subtract_datetime" => {
            crate::emitter::core::datetime_adapter::emit_datetime_subtract(chunks, current, line)
        }
        "dotnet.datetimeoffset_new" => {
            crate::emitter::core::datetime_adapter::emit_datetimeoffset_new(chunks, current, line)
        }
        "dotnet.datetimeoffset_now" | "dotnet.datetimeoffset_utc_now" => {
            crate::emitter::core::datetime_adapter::emit_datetimeoffset_utc_now(
                chunks, current, line,
            )
        }
        "dotnet.datetimeoffset_min_value" => {
            crate::emitter::core::datetime_adapter::emit_datetimeoffset_min_value(
                chunks, current, line,
            )
        }
        "dotnet.datetimeoffset_max_value" => {
            crate::emitter::core::datetime_adapter::emit_datetimeoffset_max_value(
                chunks, current, line,
            )
        }
        "dotnet.datetimeoffset_parse" => {
            crate::emitter::core::datetime_adapter::emit_datetimeoffset_parse(chunks, current, line)
        }
        "dotnet.datetimeoffset_try_parse" => {
            crate::emitter::core::datetime_adapter::emit_datetimeoffset_try_parse(
                chunks, current, line,
            )
        }
        "dotnet.datetimeoffset_from_unix_time_seconds" => {
            crate::emitter::core::datetime_adapter::emit_datetimeoffset_from_unix_time_seconds(
                chunks, current, line,
            )
        }
        "dotnet.datetimeoffset_from_unix_time_milliseconds" => {
            crate::emitter::core::datetime_adapter::emit_datetimeoffset_from_unix_time_milliseconds(
                chunks, current, line,
            )
        }
        "dotnet.datetimeoffset_add_hours" => {
            crate::emitter::core::datetime_adapter::emit_datetimeoffset_add_hours(
                chunks, current, line,
            )
        }
        "dotnet.datetimeoffset_to_universal_time" => {
            crate::emitter::core::datetime_adapter::emit_datetimeoffset_to_universal_time(
                chunks, current, line,
            )
        }
        "dotnet.datetimeoffset_to_offset" => {
            crate::emitter::core::datetime_adapter::emit_datetimeoffset_to_offset(
                chunks, current, line,
            )
        }
        "dotnet.datetimeoffset_to_unix_time_seconds" => {
            crate::emitter::core::datetime_adapter::emit_datetimeoffset_to_unix_time_seconds(
                chunks, current, line,
            )
        }
        "dotnet.datetimeoffset_to_unix_time_milliseconds" => {
            crate::emitter::core::datetime_adapter::emit_datetimeoffset_to_unix_time_milliseconds(
                chunks, current, line,
            )
        }
        "dotnet.datetimeoffset_compare" => {
            crate::emitter::core::datetime_adapter::emit_datetimeoffset_compare(
                chunks, current, line,
            )
        }
        "dotnet.datetimeoffset_subtract" => {
            crate::emitter::core::datetime_adapter::emit_datetimeoffset_subtract(
                chunks, current, line,
            )
        }
        "dotnet.datetimeoffset_equals" => {
            crate::emitter::core::datetime_adapter::emit_datetimeoffset_equals(
                chunks, current, false, line,
            )
        }
        "dotnet.datetimeoffset_equals_exact" => {
            crate::emitter::core::datetime_adapter::emit_datetimeoffset_equals(
                chunks, current, true, line,
            )
        }
        "dotnet.datetimeoffset_get_hash_code" => {
            crate::emitter::core::datetime_adapter::emit_datetimeoffset_get_hash_code(
                chunks, current, line,
            )
        }
        "dotnet.datetimeoffset_to_string" => {
            crate::emitter::core::datetime_adapter::emit_datetimeoffset_to_string(
                chunks, current, argc, line,
            )
        }

        // ── PHP DateTime / DateTimeImmutable / DateInterval adapters ──
        // Bytecode-only — composes existing `ecma:date.*` host fns into
        // the PHP-shaped surface. See `emitter/php/datetime_adapter.rs`.
        "dotnet.string_format" => crate::emitter::core::string_format_adapter::emit_string_format(
            chunks, current, argc, line,
        ),
        "dotnet.string_from_chars" => crate::emitter::core::string_adapter::emit_string_from_chars(
            chunks, current, argc, line,
        ),
        "dotnet.json_serialize" => {
            crate::emitter::core::json_adapter::emit_json_serialize(chunks, current, argc, line)
        }
        "dotnet.json_deserialize" => {
            crate::emitter::core::json_adapter::emit_json_deserialize(chunks, current, argc, line)
        }

        // ── Microsoft.VisualBasic runtime helpers shared by .NET languages ──
        "dotnet.vb_filecopy" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_filecopy(chunks, current, argc, line)
        }
        "dotnet.vb_kill" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_kill(chunks, current, argc, line)
        }
        "dotnet.vb_fileexists" => crate::emitter::core::visualbasic_adapter::emit_vb_fileexists(
            chunks, current, argc, line,
        ),
        "dotnet.vb_filelen" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_filelen(chunks, current, argc, line)
        }
        "dotnet.vb_freefile" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_freefile(chunks, current, argc, line)
        }
        "dotnet.vb_fileopen" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_fileopen(chunks, current, argc, line)
        }
        "dotnet.vb_fileclose" => crate::emitter::core::visualbasic_adapter::emit_vb_fileclose(
            chunks, current, argc, line,
        ),
        "dotnet.vb_printline" => crate::emitter::core::visualbasic_adapter::emit_vb_printline(
            chunks, current, argc, line,
        ),
        "dotnet.vb_writeline" => crate::emitter::core::visualbasic_adapter::emit_vb_writeline(
            chunks, current, argc, line,
        ),
        "dotnet.vb_lineinput" => crate::emitter::core::visualbasic_adapter::emit_vb_lineinput(
            chunks, current, argc, line,
        ),
        "dotnet.vb_input_value" => crate::emitter::core::visualbasic_adapter::emit_vb_input_value(
            chunks, current, argc, line,
        ),
        "dotnet.vb_loc" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_loc(chunks, current, argc, line)
        }
        "dotnet.vb_fileattr" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_fileattr(chunks, current, argc, line)
        }
        "dotnet.vb_getattr" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_getattr(chunks, current, argc, line)
        }
        "dotnet.vb_setattr" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_setattr(chunks, current, argc, line)
        }
        "dotnet.vb_seek" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_seek(chunks, current, argc, line)
        }
        "dotnet.vb_curdir" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_curdir(chunks, current, argc, line)
        }
        "dotnet.vb_chdir" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_chdir(chunks, current, argc, line)
        }
        "dotnet.vb_mkdir" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_mkdir(chunks, current, argc, line)
        }
        "dotnet.vb_rmdir" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_rmdir(chunks, current, argc, line)
        }
        "dotnet.vb_name" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_name(chunks, current, argc, line)
        }
        "dotnet.vb_get" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_get(chunks, current, argc, line)
        }
        "dotnet.vb_put" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_put(chunks, current, argc, line)
        }
        "dotnet.vb_app_path" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_app_path(chunks, current, argc, line)
        }
        "dotnet.vb_app_title" => crate::emitter::core::visualbasic_adapter::emit_vb_app_title(
            chunks, current, argc, line,
        ),
        "dotnet.vb_to_number" => crate::emitter::core::visualbasic_adapter::emit_vb_to_number(
            chunks, current, argc, line,
        ),
        "dotnet.vb_to_string" => crate::emitter::core::visualbasic_adapter::emit_vb_to_string(
            chunks, current, argc, line,
        ),
        "dotnet.vb_random" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_random(chunks, current, argc, line)
        }
        "dotnet.vb_lset" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_lset(chunks, current, argc, line)
        }
        "dotnet.vb_rset" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_rset(chunks, current, argc, line)
        }
        "dotnet.vb_array" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_array(chunks, current, argc, line)
        }
        "dotnet.vb_debug_print" => crate::emitter::core::visualbasic_adapter::emit_vb_debug_print(
            chunks, current, argc, line,
        ),
        "dotnet.vb_print" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_print(chunks, current, argc, line)
        }
        "dotnet.vb_input" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_input(chunks, current, argc, line)
        }
        "dotnet.vb_app" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_app(chunks, current, argc, line)
        }
        "dotnet.vb_open" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_open(chunks, current, argc, line)
        }
        "dotnet.vb_dir" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_dir(chunks, current, argc, line)
        }
        "dotnet.vb_filedatetime" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_filedatetime(
                chunks, current, argc, line,
            )
        }
        "dotnet.vb_lof" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_lof(chunks, current, argc, line)
        }
        "dotnet.vb_eof" => {
            crate::emitter::core::visualbasic_adapter::emit_vb_eof(chunks, current, argc, line)
        }
        "dotnet.vb_shell_pid" => crate::emitter::core::visualbasic_adapter::emit_vb_shell_pid(
            chunks, current, argc, line,
        ),

        // ── Microsoft.VisualBasic.Financial shared by .NET languages ────────
        "dotnet.vb_pmt" => {
            crate::emitter::core::financial_adapter::emit_vb_pmt(chunks, current, argc, line)
        }
        "dotnet.vb_fv" => {
            crate::emitter::core::financial_adapter::emit_vb_fv(chunks, current, argc, line)
        }
        "dotnet.vb_pv" => {
            crate::emitter::core::financial_adapter::emit_vb_pv(chunks, current, argc, line)
        }
        "dotnet.vb_nper" => {
            crate::emitter::core::financial_adapter::emit_vb_nper(chunks, current, argc, line)
        }
        "dotnet.vb_rate" => {
            crate::emitter::core::financial_adapter::emit_vb_rate(chunks, current, argc, line)
        }
        "dotnet.vb_ipmt" => {
            crate::emitter::core::financial_adapter::emit_vb_ipmt(chunks, current, argc, line)
        }
        "dotnet.vb_ppmt" => {
            crate::emitter::core::financial_adapter::emit_vb_ppmt(chunks, current, argc, line)
        }
        "dotnet.vb_sln" => {
            crate::emitter::core::financial_adapter::emit_vb_sln(chunks, current, argc, line)
        }
        "dotnet.vb_ddb" => {
            crate::emitter::core::financial_adapter::emit_vb_ddb(chunks, current, argc, line)
        }
        "dotnet.vb_syd" => {
            crate::emitter::core::financial_adapter::emit_vb_syd(chunks, current, argc, line)
        }

        // ── VB / VBA `Format(value, picture)` — picture-string render ──
        "dotnet.format_picture" => {
            crate::emitter::core::format_picture_adapter::emit_format_picture(
                chunks, current, argc, line,
            )
        }

        // ── .NET StreamReader / StreamWriter adapters — text I/O ────
        // Load-whole-file model: `new StreamReader(path)` materializes a
        // string buffer via `node:fs.readFileSync`, `new StreamWriter`
        // accumulates into `__buf` and flushes via `writeFileSync`.
        // Bytecode-only — no `dotnet:io` host fns.
        "dotnet.stream_reader_new" => {
            crate::emitter::core::stream_io_adapter::emit_stream_reader_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.string_reader_new" => {
            crate::emitter::core::stream_io_adapter::emit_string_reader_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.string_reader_peek" => {
            crate::emitter::core::stream_io_adapter::emit_string_reader_peek(chunks, current, line)
        }
        "dotnet.string_reader_read" => {
            crate::emitter::core::stream_io_adapter::emit_string_reader_read(chunks, current, line)
        }
        "dotnet.string_reader_read_buffer" => {
            crate::emitter::core::stream_io_adapter::emit_string_reader_read_buffer(
                chunks, current, line,
            )
        }
        "dotnet.stream_reader_read_line" => {
            crate::emitter::core::stream_io_adapter::emit_stream_reader_read_line(
                chunks, current, line,
            )
        }
        "dotnet.stream_reader_read_to_end" => {
            crate::emitter::core::stream_io_adapter::emit_stream_reader_read_to_end(
                chunks, current, line,
            )
        }
        "dotnet.stream_reader_at_end" => {
            crate::emitter::core::stream_io_adapter::emit_stream_reader_at_end(
                chunks, current, line,
            )
        }
        "dotnet.stream_reader_close" => {
            crate::emitter::core::stream_io_adapter::emit_stream_reader_close(chunks, current, line)
        }
        "dotnet.stream_writer_new" => {
            crate::emitter::core::stream_io_adapter::emit_stream_writer_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.string_writer_new" => {
            crate::emitter::core::stream_io_adapter::emit_string_writer_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.stream_writer_write" => {
            crate::emitter::core::stream_io_adapter::emit_stream_writer_write(chunks, current, line)
        }
        "dotnet.stream_writer_write_3" => {
            crate::emitter::core::stream_io_adapter::emit_stream_writer_write_3(
                chunks, current, line,
            )
        }
        "dotnet.stream_writer_write_line" => {
            crate::emitter::core::stream_io_adapter::emit_stream_writer_write_line(
                chunks, current, line,
            )
        }
        "dotnet.stream_writer_write_line_async" => {
            crate::emitter::core::stream_io_adapter::emit_stream_writer_write_line_async(
                chunks, current, line,
            )
        }
        "dotnet.stream_writer_flush" => {
            crate::emitter::core::stream_io_adapter::emit_stream_writer_flush(chunks, current, line)
        }
        "dotnet.string_writer_to_string" => {
            crate::emitter::core::stream_io_adapter::emit_string_writer_to_string(
                chunks, current, line,
            )
        }
        "dotnet.string_writer_get_string_builder" => {
            crate::emitter::core::stream_io_adapter::emit_string_writer_get_string_builder(
                chunks, current, line,
            )
        }
        "dotnet.string_writer_noop" => {
            crate::emitter::core::stream_io_adapter::emit_string_writer_noop(chunks, current, line)
        }
        "dotnet.stream_close" => {
            crate::emitter::core::stream_io_adapter::emit_stream_close(chunks, current, line)
        }
        "dotnet.file_read_all_lines" => {
            crate::emitter::core::filesystem_adapter::emit_file_read_all_lines(
                chunks, current, line,
            )
        }
        "dotnet.file_write_all_lines" => {
            crate::emitter::core::filesystem_adapter::emit_file_write_all_lines(
                chunks, current, line,
            )
        }
        "dotnet.file_read_all_bytes" => {
            crate::emitter::core::filesystem_adapter::emit_file_read_all_bytes(
                chunks, current, line,
            )
        }
        "dotnet.file_write_all_bytes" => {
            crate::emitter::core::filesystem_adapter::emit_file_write_all_bytes(
                chunks, current, line,
            )
        }
        "dotnet.file_create" => {
            crate::emitter::core::filesystem_adapter::emit_file_create(chunks, current, line)
        }
        "dotnet.file_open_read" => {
            crate::emitter::core::filesystem_adapter::emit_file_open_read(chunks, current, line)
        }
        "dotnet.file_stream_write_byte" => {
            crate::emitter::core::filesystem_adapter::emit_file_stream_write_byte(
                chunks, current, line,
            )
        }
        "dotnet.file_info_new" => {
            crate::emitter::core::filesystem_adapter::emit_file_info_new(chunks, current, line)
        }
        "dotnet.path_combine" => {
            crate::emitter::core::filesystem_adapter::emit_path_combine(chunks, current, argc, line)
        }
        "dotnet.path_get_file_name" => {
            crate::emitter::core::filesystem_adapter::emit_path_get_file_name(chunks, current, line)
        }
        "dotnet.path_get_directory_name" => {
            crate::emitter::core::filesystem_adapter::emit_path_get_directory_name(
                chunks, current, line,
            )
        }
        "dotnet.path_get_file_name_without_extension" => {
            crate::emitter::core::filesystem_adapter::emit_path_get_file_name_without_extension(
                chunks, current, line,
            )
        }
        "dotnet.path_change_extension" => {
            crate::emitter::core::filesystem_adapter::emit_path_change_extension(
                chunks, current, line,
            )
        }
        "dotnet.path_get_full_path" => {
            crate::emitter::core::filesystem_adapter::emit_path_get_full_path(chunks, current, line)
        }
        "dotnet.path_get_path_root" => {
            crate::emitter::core::filesystem_adapter::emit_path_get_path_root(chunks, current, line)
        }
        "dotnet.path_get_temp_file_name" => {
            crate::emitter::core::filesystem_adapter::emit_path_get_temp_file_name(
                chunks, current, line,
            )
        }
        "dotnet.path_get_random_file_name" => {
            crate::emitter::core::filesystem_adapter::emit_path_get_random_file_name(
                chunks, current, line,
            )
        }
        "dotnet.path_get_invalid_file_name_chars" => {
            crate::emitter::core::filesystem_adapter::emit_path_get_invalid_file_name_chars(
                chunks, current, line,
            )
        }
        "dotnet.path_get_invalid_path_chars" => {
            crate::emitter::core::filesystem_adapter::emit_path_get_invalid_path_chars(
                chunks, current, line,
            )
        }
        "dotnet.path_has_extension" => {
            crate::emitter::core::filesystem_adapter::emit_path_has_extension(chunks, current, line)
        }
        "dotnet.path_is_path_rooted" => {
            crate::emitter::core::filesystem_adapter::emit_path_is_path_rooted(
                chunks, current, line,
            )
        }
        "dotnet.path_get_relative_path" => {
            crate::emitter::core::filesystem_adapter::emit_path_get_relative_path(
                chunks, current, line,
            )
        }
        "dotnet.path_trim_ending_directory_separator" => {
            crate::emitter::core::filesystem_adapter::emit_path_trim_ending_directory_separator(
                chunks, current, line,
            )
        }
        "dotnet.directory_get_files" => {
            crate::emitter::core::filesystem_adapter::emit_directory_get_files(
                chunks, current, argc, line,
            )
        }
        "dotnet.directory_get_directories" => {
            crate::emitter::core::filesystem_adapter::emit_directory_get_directories(
                chunks, current, line,
            )
        }
        "dotnet.directory_delete" => {
            crate::emitter::core::filesystem_adapter::emit_directory_delete(
                chunks, current, argc, line,
            )
        }
        "dotnet.console_writeline" => {
            // `WriteLine()` with no argument is a bare newline; `WriteLine(v)`
            // stringifies then appends one. The call site passes the real argc.
            if argc == 0 {
                crate::emitter::core::console_adapter::emit_console_writeline_empty(
                    chunks, current, line,
                )
            } else {
                crate::emitter::core::console_adapter::emit_console_writeline(chunks, current, line)
            }
        }
        "dotnet.console_write" => {
            crate::emitter::core::console_adapter::emit_console_write(chunks, current, line)
        }
        "dotnet.console_readline" => {
            crate::emitter::core::console_adapter::emit_console_readline(chunks, current, line)
        }
        "dotnet.console_error" => {
            crate::emitter::core::console_adapter::emit_console_error(chunks, current, line)
        }
        "dotnet.console_error_write" => {
            crate::emitter::core::console_adapter::emit_console_error_write(chunks, current, line)
        }
        "dotnet.console_error_writeline" => {
            crate::emitter::core::console_adapter::emit_console_error_writeline(
                chunks, current, line,
            )
        }
        "dotnet.encoding_utf8" => crate::emitter::core::encoding_adapter::emit_encoding_value(
            chunks, current, "utf-8", line,
        ),
        "dotnet.encoding_ascii" => crate::emitter::core::encoding_adapter::emit_encoding_value(
            chunks, current, "ascii", line,
        ),
        "dotnet.encoding_unicode" => crate::emitter::core::encoding_adapter::emit_encoding_value(
            chunks, current, "utf16le", line,
        ),
        "dotnet.encoding_utf32" => crate::emitter::core::encoding_adapter::emit_encoding_value(
            chunks, current, "utf32", line,
        ),
        "dotnet.encoding_big_endian_unicode" => {
            crate::emitter::core::encoding_adapter::emit_encoding_value(
                chunks, current, "utf16be", line,
            )
        }
        "dotnet.encoding_latin1" => crate::emitter::core::encoding_adapter::emit_encoding_value(
            chunks, current, "latin1", line,
        ),
        "dotnet.utf8encoding_new" => {
            for _ in 0..argc {
                chunks[current].emit_op(vybe_runtime::opcode::Op::DROP, line);
            }
            crate::emitter::core::encoding_adapter::emit_encoding_value(
                chunks, current, "utf-8", line,
            )
        }
        "dotnet.encoding_get_encoding" => {
            crate::emitter::core::encoding_adapter::emit_encoding_get_encoding(
                chunks, current, argc, line,
            )
        }
        "dotnet.encoding_convert" => {
            crate::emitter::core::encoding_adapter::emit_encoding_convert(chunks, current, line)
        }
        "dotnet.encoding_get_bytes" => {
            crate::emitter::core::encoding_adapter::emit_encoding_get_bytes(
                chunks, current, argc, "utf8", line,
            )
        }
        "dotnet.encoding_ascii_get_bytes" => {
            crate::emitter::core::encoding_adapter::emit_encoding_get_bytes(
                chunks, current, argc, "ascii", line,
            )
        }
        "dotnet.encoding_unicode_get_bytes" => {
            crate::emitter::core::encoding_adapter::emit_encoding_get_bytes(
                chunks, current, argc, "utf16le", line,
            )
        }
        "dotnet.encoding_utf32_get_bytes" => {
            crate::emitter::core::encoding_adapter::emit_encoding_utf32_get_bytes(
                chunks, current, argc, line,
            )
        }
        "dotnet.encoding_big_endian_unicode_get_bytes" => {
            crate::emitter::core::encoding_adapter::emit_encoding_big_endian_unicode_get_bytes(
                chunks, current, argc, line,
            )
        }
        "dotnet.encoding_latin1_get_bytes" => {
            crate::emitter::core::encoding_adapter::emit_encoding_get_bytes(
                chunks, current, argc, "latin1", line,
            )
        }
        "dotnet.encoding_get_string" => {
            crate::emitter::core::encoding_adapter::emit_encoding_get_string(
                chunks, current, argc, "utf8", line,
            )
        }
        "dotnet.encoding_ascii_get_string" => {
            crate::emitter::core::encoding_adapter::emit_encoding_get_string(
                chunks, current, argc, "ascii", line,
            )
        }
        "dotnet.encoding_unicode_get_string" => {
            crate::emitter::core::encoding_adapter::emit_encoding_unicode_get_string(
                chunks, current, argc, line,
            )
        }
        "dotnet.encoding_utf32_get_string" => {
            crate::emitter::core::encoding_adapter::emit_encoding_utf32_get_string(
                chunks, current, argc, line,
            )
        }
        "dotnet.encoding_latin1_get_string" => {
            crate::emitter::core::encoding_adapter::emit_encoding_get_string(
                chunks, current, argc, "latin1", line,
            )
        }
        "dotnet.encoding_get_byte_count" => {
            crate::emitter::core::encoding_adapter::emit_encoding_get_byte_count(
                chunks, current, argc, "utf8", line,
            )
        }
        "dotnet.encoding_ascii_get_byte_count" => {
            crate::emitter::core::encoding_adapter::emit_encoding_get_byte_count(
                chunks, current, argc, "ascii", line,
            )
        }
        "dotnet.encoding_unicode_get_byte_count" => {
            crate::emitter::core::encoding_adapter::emit_encoding_get_byte_count(
                chunks, current, argc, "utf16le", line,
            )
        }
        "dotnet.encoding_utf32_get_byte_count" => {
            crate::emitter::core::encoding_adapter::emit_encoding_utf32_get_byte_count(
                chunks, current, argc, line,
            )
        }
        "dotnet.encoding_latin1_get_byte_count" => {
            crate::emitter::core::encoding_adapter::emit_encoding_get_byte_count(
                chunks, current, argc, "latin1", line,
            )
        }
        "dotnet.encoding_get_preamble" => {
            crate::emitter::core::encoding_adapter::emit_encoding_get_preamble(
                chunks, current, argc, line,
            )
        }
        "dotnet.encoding_utf8_get_preamble" => {
            crate::emitter::core::encoding_adapter::emit_encoding_utf8_get_preamble(
                chunks, current, argc, line,
            )
        }
        "dotnet.encoding_get_max_byte_count" => {
            crate::emitter::core::encoding_adapter::emit_encoding_get_max_byte_count(
                chunks, current, argc, 4, line,
            )
        }
        "dotnet.encoding_unicode_get_max_byte_count" => {
            crate::emitter::core::encoding_adapter::emit_encoding_get_max_byte_count(
                chunks, current, argc, 2, line,
            )
        }
        "dotnet.encoding_get_max_char_count" => {
            crate::emitter::core::encoding_adapter::emit_encoding_get_max_char_count(
                chunks, current, argc, line,
            )
        }
        "dotnet.encoding_get_char_count" => {
            crate::emitter::core::encoding_adapter::emit_encoding_get_char_count(
                chunks, current, argc, line,
            )
        }
        "dotnet.encoding_get_chars" => {
            crate::emitter::core::encoding_adapter::emit_encoding_get_chars(
                chunks, current, argc, line,
            )
        }
        "dotnet.encoding_equals" => {
            crate::emitter::core::encoding_adapter::emit_encoding_equals(chunks, current, line)
        }
        "dotnet.object_equals" => {
            crate::emitter::core::encoding_adapter::emit_object_equals(chunks, current, line)
        }
        "dotnet.object_reference_equals" => {
            chunks[current].emit_op(Op::REF_EQ, line);
            vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "dotnet.object_to_string_role" => {
            let slot = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
            vybe_compiler::primitives::expressions::emit_rich_to_string(
                &mut chunks[current],
                slot,
                line,
            );
        }
        "dotnet.gc_noop" => {
            crate::emitter::core::gc_adapter::emit_gc_noop(chunks, current, argc, line)
        }
        "dotnet.gc_zero" => {
            crate::emitter::core::gc_adapter::emit_gc_zero(chunks, current, argc, line)
        }
        "dotnet.gc_total_memory" => {
            crate::emitter::core::gc_adapter::emit_gc_get_total_memory(chunks, current, argc, line)
        }
        "dotnet.gc_generation" => {
            crate::emitter::core::gc_adapter::emit_gc_get_generation(chunks, current, line)
        }
        "dotnet.identity" => {}
        "dotnet.environment_username" => {
            crate::emitter::core::environment_adapter::emit_environment_username(
                chunks, current, line,
            )
        }
        "dotnet.environment_version" => {
            crate::emitter::core::environment_adapter::emit_environment_version(
                chunks, current, line,
            )
        }
        "dotnet.environment_exit_code" => {
            crate::emitter::core::environment_adapter::emit_environment_exit_code(
                chunks, current, line,
            )
        }
        "dotnet.environment_set_exit_code" => {
            crate::emitter::core::environment_adapter::emit_environment_set_exit_code(
                chunks, current, line,
            )
        }
        "dotnet.environment_system_directory" => {
            crate::emitter::core::environment_adapter::emit_environment_system_directory(
                chunks, current, line,
            )
        }
        "dotnet.environment_processor_count" => {
            crate::emitter::core::environment_adapter::emit_environment_processor_count(
                chunks, current, line,
            )
        }
        "dotnet.environment_tick_count" => {
            crate::emitter::core::environment_adapter::emit_environment_tick_count(
                chunks, current, line,
            )
        }
        "dotnet.environment_get" => {
            crate::emitter::core::environment_adapter::emit_environment_get(
                chunks, current, argc, line,
            )
        }
        "dotnet.environment_set" => {
            crate::emitter::core::environment_adapter::emit_environment_set(
                chunks, current, argc, line,
            )
        }
        "dotnet.environment_get_all" => {
            crate::emitter::core::environment_adapter::emit_environment_get_all(
                chunks, current, argc, line,
            )
        }
        "dotnet.environment_expand" => {
            crate::emitter::core::environment_adapter::emit_environment_expand(
                chunks, current, line,
            )
        }
        "dotnet.environment_target_process" => {
            crate::emitter::core::environment_adapter::emit_environment_target(
                "Process", chunks, current, line,
            )
        }
        "dotnet.environment_target_user" => {
            crate::emitter::core::environment_adapter::emit_environment_target(
                "User", chunks, current, line,
            )
        }
        "dotnet.environment_target_machine" => {
            crate::emitter::core::environment_adapter::emit_environment_target(
                "Machine", chunks, current, line,
            )
        }
        "dotnet.environment_get_folder_path" => {
            crate::emitter::core::environment_adapter::emit_environment_get_folder_path(
                chunks, current, line,
            )
        }
        "dotnet.environment_get_command_line_args" => {
            crate::emitter::core::environment_adapter::emit_environment_get_command_line_args(
                chunks, current, line,
            )
        }
        "dotnet.environment_special_folder_personal" => {
            crate::emitter::core::environment_adapter::emit_environment_special_folder(
                "Personal", chunks, current, line,
            )
        }
        "dotnet.environment_special_folder_application_data" => {
            crate::emitter::core::environment_adapter::emit_environment_special_folder(
                "ApplicationData",
                chunks,
                current,
                line,
            )
        }
        "dotnet.environment_special_folder_local_application_data" => {
            crate::emitter::core::environment_adapter::emit_environment_special_folder(
                "LocalApplicationData",
                chunks,
                current,
                line,
            )
        }
        "dotnet.environment_special_folder_desktop" => {
            crate::emitter::core::environment_adapter::emit_environment_special_folder(
                "Desktop", chunks, current, line,
            )
        }

        // ── OleDb adapter — System.Data.OleDb constructor wrappers ─────────────
        "dotnet.oledb_connection_new" => {
            crate::emitter::core::oledb_adapter::emit_oledb_connection_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.oledb_command_new" => {
            crate::emitter::core::oledb_adapter::emit_oledb_command_new(chunks, current, argc, line)
        }

        // ── ADODB adapter — ADODB.Connection / Command / Recordset ──────────────
        "dotnet.adodb_connection_new" => {
            crate::emitter::core::adodb_adapter::emit_adodb_connection_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_connection_execute" => {
            crate::emitter::core::adodb_adapter::emit_adodb_connection_execute(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_conn_begin_trans" => {
            crate::emitter::core::adodb_adapter::emit_adodb_conn_begin_trans(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_conn_commit_trans" => {
            crate::emitter::core::adodb_adapter::emit_adodb_conn_commit_trans(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_conn_rollback_trans" => {
            crate::emitter::core::adodb_adapter::emit_adodb_conn_rollback_trans(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_command_new" => {
            crate::emitter::core::adodb_adapter::emit_adodb_command_new(chunks, current, argc, line)
        }
        "dotnet.adodb_command_execute" => {
            crate::emitter::core::adodb_adapter::emit_adodb_command_execute(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_command_create_parameter" => {
            crate::emitter::core::adodb_adapter::emit_adodb_command_create_parameter(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_recordset_new" => {
            crate::emitter::core::adodb_adapter::emit_adodb_recordset_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_recordset_open" => {
            crate::emitter::core::adodb_adapter::emit_adodb_recordset_open(
                chunks, current, argc, line,
            )
        }
        "dotnet.adodb_recordset_move_next" => {
            crate::emitter::core::adodb_adapter::emit_adodb_recordset_move_next(
                chunks, current, line,
            )
        }
        "dotnet.adodb_recordset_move_first" => {
            crate::emitter::core::adodb_adapter::emit_adodb_recordset_move_first(
                chunks, current, line,
            )
        }
        "dotnet.adodb_recordset_fields" => {
            crate::emitter::core::adodb_adapter::emit_adodb_recordset_fields(chunks, current, line)
        }
        "dotnet.adodb_recordset_close" => {
            crate::emitter::core::adodb_adapter::emit_adodb_recordset_close(chunks, current, line)
        }

        // ── LINQ surface — composed bytecode shared by every .NET-shape language ──
        "dotnet.linq_first" => {
            crate::emitter::core::linq_adapter::emit_linq_first(chunks, current, line)
        }
        "dotnet.linq_last" => {
            crate::emitter::core::linq_adapter::emit_linq_last(chunks, current, line)
        }
        "dotnet.linq_last_or_default" => {
            crate::emitter::core::linq_adapter::emit_linq_last_or_default(chunks, current, line)
        }
        "dotnet.linq_skip" => {
            crate::emitter::core::linq_adapter::emit_linq_skip(chunks, current, line)
        }
        "dotnet.linq_take" => {
            crate::emitter::core::linq_adapter::emit_linq_take(chunks, current, line)
        }
        "dotnet.linq_identity" => {
            crate::emitter::core::linq_adapter::emit_linq_identity(chunks, current, line)
        }
        "dotnet.linq_average" => {
            crate::emitter::core::linq_adapter::emit_linq_average(chunks, current, line)
        }
        "dotnet.linq_first_or_default" => {
            crate::emitter::core::linq_adapter::emit_linq_first_or_default(chunks, current, line)
        }
        "dotnet.linq_distinct" => {
            crate::emitter::core::linq_adapter::emit_linq_distinct(chunks, current, line)
        }
        "dotnet.linq_distinct_comparer" => {
            crate::emitter::core::linq_adapter::emit_linq_distinct_comparer(chunks, current, line)
        }
        "dotnet.linq_distinct_by" => {
            crate::emitter::core::linq_adapter::emit_linq_distinct_by(chunks, current, line)
        }
        "dotnet.linq_distinct_by_comparer" => {
            crate::emitter::core::linq_adapter::emit_linq_distinct_by_comparer(
                chunks, current, line,
            )
        }
        "dotnet.linq_order_by" => {
            crate::emitter::core::linq_adapter::emit_linq_order_by(chunks, current, line)
        }
        "dotnet.linq_sequence_equal" => {
            crate::emitter::core::linq_adapter::emit_linq_sequence_equal(chunks, current, line)
        }
        "dotnet.linq_sequence_equal_comparer" => {
            crate::emitter::core::linq_adapter::emit_linq_sequence_equal_comparer(
                chunks, current, line,
            )
        }
        "dotnet.linq_all" => {
            crate::emitter::core::linq_adapter::emit_linq_all(chunks, current, line)
        }
        "dotnet.linq_count_pred" => {
            crate::emitter::core::linq_adapter::emit_linq_count_pred(chunks, current, line)
        }
        "dotnet.linq_aggregate" => {
            crate::emitter::core::linq_adapter::emit_linq_aggregate(chunks, current, line)
        }
        "dotnet.linq_order_by_descending" => {
            crate::emitter::core::linq_adapter::emit_linq_order_by_descending(chunks, current, line)
        }
        "dotnet.linq_select" => {
            crate::emitter::core::linq_adapter::emit_linq_select(chunks, current, line)
        }
        "dotnet.linq_select_many" => {
            crate::emitter::core::linq_adapter::emit_linq_select_many(chunks, current, line)
        }
        "dotnet.linq_group_by" => {
            crate::emitter::core::linq_adapter::emit_linq_group_by(chunks, current, line)
        }
        "dotnet.linq_to_dictionary" => {
            crate::emitter::core::linq_adapter::emit_linq_to_dictionary(chunks, current, line)
        }
        "dotnet.linq_to_dictionary_key" => {
            crate::emitter::core::linq_adapter::emit_linq_to_dictionary_key(chunks, current, line)
        }
        "dotnet.linq_to_lookup" => {
            crate::emitter::core::linq_adapter::emit_linq_to_lookup(chunks, current, line)
        }
        "dotnet.linq_zip" => {
            crate::emitter::core::linq_adapter::emit_linq_zip(chunks, current, line)
        }
        "dotnet.linq_concat" => {
            crate::emitter::core::linq_adapter::emit_linq_concat(chunks, current, line)
        }
        "dotnet.linq_union" => {
            crate::emitter::core::linq_adapter::emit_linq_union(chunks, current, line)
        }
        "dotnet.linq_union_comparer" => {
            crate::emitter::core::linq_adapter::emit_linq_union_comparer(chunks, current, line)
        }
        "dotnet.linq_union_by" => {
            crate::emitter::core::linq_adapter::emit_linq_union_by(chunks, current, line)
        }
        "dotnet.linq_intersect" => {
            crate::emitter::core::linq_adapter::emit_linq_intersect(chunks, current, line)
        }
        "dotnet.linq_intersect_comparer" => {
            crate::emitter::core::linq_adapter::emit_linq_intersect_comparer(chunks, current, line)
        }
        "dotnet.linq_intersect_by" => {
            crate::emitter::core::linq_adapter::emit_linq_intersect_by(chunks, current, line)
        }
        "dotnet.linq_except" => {
            crate::emitter::core::linq_adapter::emit_linq_except(chunks, current, line)
        }
        "dotnet.linq_except_comparer" => {
            crate::emitter::core::linq_adapter::emit_linq_except_comparer(chunks, current, line)
        }
        "dotnet.linq_except_by" => {
            crate::emitter::core::linq_adapter::emit_linq_except_by(chunks, current, line)
        }
        "dotnet.linq_of_type" => {
            crate::emitter::core::linq_adapter::emit_linq_of_type(chunks, current, line)
        }
        "dotnet.linq_element_at" => {
            crate::emitter::core::linq_adapter::emit_linq_element_at(chunks, current, line)
        }
        "dotnet.linq_element_at_or_default" => {
            crate::emitter::core::linq_adapter::emit_linq_element_at_or_default(
                chunks, current, line,
            )
        }
        "dotnet.linq_single" => {
            crate::emitter::core::linq_adapter::emit_linq_single(chunks, current, line)
        }
        "dotnet.linq_single_or_default" => {
            crate::emitter::core::linq_adapter::emit_linq_single_or_default(chunks, current, line)
        }
        "dotnet.linq_max_by" => {
            crate::emitter::core::linq_adapter::emit_linq_max_by(chunks, current, line)
        }
        "dotnet.linq_min_by" => {
            crate::emitter::core::linq_adapter::emit_linq_min_by(chunks, current, line)
        }
        "dotnet.linq_aggregate_no_seed" => {
            crate::emitter::core::linq_adapter::emit_linq_aggregate_no_seed(chunks, current, line)
        }
        "dotnet.linq_append" => {
            crate::emitter::core::linq_adapter::emit_linq_append(chunks, current, line)
        }
        "dotnet.linq_prepend" => {
            crate::emitter::core::linq_adapter::emit_linq_prepend(chunks, current, line)
        }
        "dotnet.linq_sum" => {
            crate::emitter::core::linq_adapter::emit_linq_sum(chunks, current, line)
        }
        "dotnet.linq_sum_selector" => {
            crate::emitter::core::linq_adapter::emit_linq_sum_selector(chunks, current, line)
        }
        "dotnet.linq_count" => {
            crate::emitter::core::linq_adapter::emit_linq_count(chunks, current, line)
        }
        "dotnet.linq_where" => {
            crate::emitter::core::linq_adapter::emit_linq_where(chunks, current, line)
        }
        "dotnet.linq_any" => {
            crate::emitter::core::linq_adapter::emit_linq_any(chunks, current, line)
        }
        "dotnet.linq_any_pred" => {
            crate::emitter::core::linq_adapter::emit_linq_any_pred(chunks, current, line)
        }
        "dotnet.linq_contains" => {
            crate::emitter::core::linq_adapter::emit_linq_contains(chunks, current, line)
        }
        "dotnet.linq_contains_comparer" => {
            crate::emitter::core::linq_adapter::emit_linq_contains_comparer(chunks, current, line)
        }
        "dotnet.linq_reverse" => {
            crate::emitter::core::linq_adapter::emit_linq_reverse(chunks, current, line)
        }
        "dotnet.linq_skip_while" => {
            crate::emitter::core::linq_adapter::emit_linq_skip_while(chunks, current, line)
        }
        "dotnet.linq_skip_while_indexed" => {
            crate::emitter::core::linq_adapter::emit_linq_skip_while_indexed(chunks, current, line)
        }
        "dotnet.linq_take_while" => {
            crate::emitter::core::linq_adapter::emit_linq_take_while(chunks, current, line)
        }
        "dotnet.linq_take_while_indexed" => {
            crate::emitter::core::linq_adapter::emit_linq_take_while_indexed(chunks, current, line)
        }
        "dotnet.linq_chunk" => {
            crate::emitter::core::linq_adapter::emit_linq_chunk(chunks, current, line)
        }
        "dotnet.linq_skip_last" => {
            crate::emitter::core::linq_adapter::emit_linq_skip_last(chunks, current, line)
        }
        "dotnet.linq_take_last" => {
            crate::emitter::core::linq_adapter::emit_linq_take_last(chunks, current, line)
        }
        "dotnet.linq_default_if_empty" => {
            crate::emitter::core::linq_adapter::emit_linq_default_if_empty(chunks, current, line)
        }
        "dotnet.linq_default_if_empty_value" => {
            crate::emitter::core::linq_adapter::emit_linq_default_if_empty_value(
                chunks, current, line,
            )
        }

        // ── Static Array.* helpers — same dotnet/core home as LINQ ──
        "dotnet.array_reverse" => crate::emitter::core::array_adapter::emit_array_reverse_arity(
            chunks, current, argc, line,
        ),
        "dotnet.array_fill" => {
            crate::emitter::core::array_adapter::emit_array_fill(chunks, current, argc, line)
        }
        "dotnet.array_index_of" => crate::emitter::core::array_adapter::emit_array_index_of_arity(
            chunks, current, argc, line,
        ),
        "dotnet.array_last_index_of" => {
            crate::emitter::core::array_adapter::emit_array_last_index_of(
                chunks, current, argc, line,
            )
        }
        "dotnet.array_exists" => {
            crate::emitter::core::array_adapter::emit_array_exists(chunks, current, line)
        }
        "dotnet.array_true_for_all" => {
            crate::emitter::core::array_adapter::emit_array_true_for_all(chunks, current, line)
        }
        "dotnet.array_find" => {
            crate::emitter::core::array_adapter::emit_array_find(chunks, current, line)
        }
        "dotnet.array_find_last" => {
            crate::emitter::core::array_adapter::emit_array_find_last(chunks, current, line)
        }
        "dotnet.array_find_all" => {
            crate::emitter::core::array_adapter::emit_array_find_all(chunks, current, line)
        }
        "dotnet.array_find_index" => {
            crate::emitter::core::array_adapter::emit_array_find_index(chunks, current, argc, line)
        }
        "dotnet.array_find_last_index" => {
            crate::emitter::core::array_adapter::emit_array_find_last_index(
                chunks, current, argc, line,
            )
        }
        "dotnet.array_binary_search" => {
            crate::emitter::core::array_adapter::emit_array_binary_search(
                chunks, current, argc, line,
            )
        }
        "dotnet.array_create_instance" => {
            crate::emitter::core::array_adapter::emit_array_create_instance(
                chunks, current, argc, line,
            )
        }
        "dotnet.array_empty" => {
            crate::emitter::core::array_adapter::emit_array_empty(chunks, current, line)
        }
        "dotnet.array_convert_all" => {
            crate::emitter::core::array_adapter::emit_array_convert_all(chunks, current, line)
        }
        "dotnet.array_for_each" => {
            crate::emitter::core::array_adapter::emit_array_for_each(chunks, current, line)
        }
        "dotnet.list_add_range" => {
            crate::emitter::core::array_adapter::emit_list_add_range(chunks, current, line)
        }
        "dotnet.convert_to_base64_string" => {
            crate::emitter::core::convert_adapter::emit_convert_to_base64_string(
                chunks, current, argc, line,
            )
        }
        "dotnet.convert_from_base64_string" => {
            crate::emitter::core::convert_adapter::emit_convert_from_base64_string(
                chunks, current, argc, line,
            )
        }
        "dotnet.convert_try_from_base64_chars" => {
            crate::emitter::core::convert_adapter::emit_convert_try_from_base64_chars(
                chunks, current, argc, line,
            )
        }
        "dotnet.convert_to_base64_char_array" => {
            crate::emitter::core::convert_adapter::emit_convert_to_base64_char_array(
                chunks, current, argc, line,
            )
        }

        // ── System.Span<T> members with no 1:1 ECMA array method ──────
        // (the ones that DO line up are plain `HostCall`s on the
        // descriptor — see `component_classes_span.rs`)
        "dotnet.span_ctor" => {
            crate::emitter::core::span_adapter::emit_span_ctor(chunks, current, argc, line)
        }
        "dotnet.span_is_empty" => {
            crate::emitter::core::span_adapter::emit_span_is_empty(chunks, current, line)
        }
        "dotnet.span_clear" => {
            crate::emitter::core::span_adapter::emit_span_clear(chunks, current, line)
        }
        "dotnet.span_copy_to" => {
            crate::emitter::core::span_adapter::emit_span_copy_to(chunks, current, line)
        }
        "dotnet.span_try_copy_to" => {
            crate::emitter::core::span_adapter::emit_span_try_copy_to(chunks, current, line)
        }
        "dotnet.span_trim_start" => {
            crate::emitter::core::span_adapter::emit_span_trim_start(chunks, current, line)
        }
        "dotnet.span_trim_end" => {
            crate::emitter::core::span_adapter::emit_span_trim_end(chunks, current, line)
        }
        "dotnet.span_mismatch" => {
            crate::emitter::core::span_adapter::emit_span_mismatch(chunks, current, line)
        }
        "dotnet.array_segment_ctor" => {
            crate::emitter::core::span_adapter::emit_array_segment_ctor(chunks, current, argc, line)
        }
        "dotnet.array_segment_empty" => {
            crate::emitter::core::span_adapter::emit_array_segment_empty(chunks, current, line)
        }
        "dotnet.array_segment_get" => {
            crate::emitter::core::span_adapter::emit_array_segment_get(chunks, current, line)
        }
        "dotnet.array_segment_set" => {
            crate::emitter::core::span_adapter::emit_array_segment_set(chunks, current, line)
        }
        "dotnet.array_segment_slice" => {
            crate::emitter::core::span_adapter::emit_array_segment_slice(chunks, current, line)
        }
        "dotnet.array_segment_copy_to" => {
            crate::emitter::core::span_adapter::emit_array_segment_copy_to(chunks, current, line)
        }
        "dotnet.array_segment_to_array" => {
            crate::emitter::core::span_adapter::emit_array_segment_to_array(chunks, current, line)
        }
        "dotnet.array_segment_equals" => {
            crate::emitter::core::span_adapter::emit_array_segment_equals(chunks, current, line)
        }
        "dotnet.array_pool_shared" => {
            crate::emitter::core::span_adapter::emit_array_pool_shared(chunks, current, line)
        }
        "dotnet.array_pool_rent" => {
            crate::emitter::core::span_adapter::emit_array_pool_rent(chunks, current, line)
        }
        "dotnet.array_pool_rent_static" => {
            crate::emitter::core::span_adapter::emit_array_pool_rent_static(chunks, current, line)
        }
        "dotnet.array_pool_return" => {
            crate::emitter::core::span_adapter::emit_array_pool_return(chunks, current, argc, line)
        }
        "dotnet.memory_pool_shared" => {
            crate::emitter::core::span_adapter::emit_memory_pool_shared(chunks, current, line)
        }
        "dotnet.memory_pool_rent" => {
            crate::emitter::core::span_adapter::emit_memory_pool_rent(chunks, current, line)
        }
        "dotnet.memory_pool_rent_static" => {
            crate::emitter::core::span_adapter::emit_memory_pool_rent_static(chunks, current, line)
        }
        "dotnet.buffer_byte_length" => {
            crate::emitter::core::span_adapter::emit_buffer_byte_length(chunks, current, line)
        }
        "dotnet.buffer_get_byte" => {
            crate::emitter::core::span_adapter::emit_buffer_get_byte(chunks, current, line)
        }
        "dotnet.buffer_set_byte" => {
            crate::emitter::core::span_adapter::emit_buffer_set_byte(chunks, current, line)
        }

        // ── .NET parse helpers — `int.Parse`, `double.Parse`, `bool.Parse`
        // Throw a `FormatException`-shape error on invalid input
        // (matches ECMA-335; JS `Number(s)` returns NaN silently).
        "dotnet.parse_int" => {
            crate::emitter::core::parse_adapter::emit_parse_int(chunks, current, line)
        }
        "dotnet.try_parse_int" => {
            crate::emitter::core::parse_adapter::emit_try_parse_int(chunks, current, line)
        }
        "dotnet.parse_byte" => {
            crate::emitter::core::parse_adapter::emit_parse_int(chunks, current, line)
        }
        "dotnet.parse_long" => {
            crate::emitter::core::parse_adapter::emit_parse_int(chunks, current, line)
        }
        "dotnet.parse_float" => {
            crate::emitter::core::parse_adapter::emit_parse_double(chunks, current, line)
        }
        "dotnet.parse_decimal" => {
            crate::emitter::core::parse_adapter::emit_parse_double(chunks, current, line)
        }
        "dotnet.parse_double" => {
            crate::emitter::core::parse_adapter::emit_parse_double(chunks, current, line)
        }
        "dotnet.parse_bool" => {
            crate::emitter::core::parse_adapter::emit_parse_bool(chunks, current, line)
        }
        "dotnet.parse_char" => {
            crate::emitter::core::parse_adapter::emit_parse_char(chunks, current, line)
        }

        // ── System.Windows.Forms.BindingSource — the data cursor ────
        "dotnet.bindingsource_new" => {
            crate::emitter::core::bindingsource_adapter::emit_bindingsource_new(
                chunks, current, argc, line,
            )
        }
        "dotnet.bindingsource_move_first" => {
            crate::emitter::core::bindingsource_adapter::emit_bindingsource_move(
                chunks,
                current,
                crate::emitter::core::bindingsource_adapter::Move::First,
                line,
            )
        }
        "dotnet.bindingsource_move_next" => {
            crate::emitter::core::bindingsource_adapter::emit_bindingsource_move(
                chunks,
                current,
                crate::emitter::core::bindingsource_adapter::Move::Next,
                line,
            )
        }
        "dotnet.bindingsource_move_previous" => {
            crate::emitter::core::bindingsource_adapter::emit_bindingsource_move(
                chunks,
                current,
                crate::emitter::core::bindingsource_adapter::Move::Previous,
                line,
            )
        }
        "dotnet.bindingsource_move_last" => {
            crate::emitter::core::bindingsource_adapter::emit_bindingsource_move(
                chunks,
                current,
                crate::emitter::core::bindingsource_adapter::Move::Last,
                line,
            )
        }
        "dotnet.bindingsource_count" => {
            crate::emitter::core::bindingsource_adapter::emit_bindingsource_count(
                chunks, current, line,
            )
        }
        "dotnet.bindingsource_current" => {
            crate::emitter::core::bindingsource_adapter::emit_bindingsource_current(
                chunks, current, line,
            )
        }

        // ── .NET System.Data adapter ────────────────────────────────
        "dotnet.datatable_new" => {
            crate::emitter::core::datatable_adapter::emit_datatable_new(chunks, current, argc, line)
        }
        "dotnet.dataset_new" => {
            crate::emitter::core::datatable_adapter::emit_dataset_new(chunks, current, argc, line)
        }
        "dotnet.datarow_new" => {
            crate::emitter::core::datatable_adapter::emit_datarow_new(&mut chunks[current], line)
        }
        "dotnet.datatable_new_row" => {
            crate::emitter::core::datatable_adapter::emit_datatable_new_row(chunks, current, line)
        }
        "dotnet.datatable_add_row" => {
            crate::emitter::core::datatable_adapter::emit_datatable_add_row(chunks, current, line)
        }
        "dotnet.datatable_select" => {
            crate::emitter::core::datatable_adapter::emit_datatable_select(chunks, current, line)
        }
        "dotnet.dataset_tables" => {
            crate::emitter::core::datatable_adapter::emit_dataset_tables(chunks, current, line)
        }
        "dotnet.datarow_item" => {
            crate::emitter::core::datatable_adapter::emit_datarow_item(chunks, current, line)
        }
        "dotnet.datarow_is_null" => {
            crate::emitter::core::datatable_adapter::emit_datarow_is_null(chunks, current, line)
        }

        // ── PHP `isset(...)` — variadic null check, returns true iff
        // ALL args are non-null. Inline emit folds an AND chain.
        "dotnet.dict_get_or_throw" => {
            // map[key] — get or throw KeyNotFoundException
            let chunk = &mut chunks[current];
            let has = chunk.add_import("ecma:map", "has");
            let get = chunk.add_import("ecma:map", "get");
            let key_slot = chunk.alloc_scratch(1);
            let map_slot = chunk.alloc_scratch(1);
            chunk.emit_op_u16(Op::LOCAL_SET, key_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, map_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, map_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
            chunk.emit_call(has, 2, line);
            chunk.emit_if_value(line);
            chunk.emit_op_u16(Op::LOCAL_GET, map_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
            chunk.emit_call(get, 2, line);
            chunk.emit_else(line);
            chunk.emit_struct_new(0, 0, line);
            chunk.emit_dup(line);
            chunk.emit_string_const("The given key was not present in the dictionary.", line);
            vybe_compiler::primitives::errors::emit_exception_new_finalize(
                chunk,
                "KeyNotFoundException",
                line,
            );
            vybe_compiler::primitives::errors::emit_stamp_exception_ancestors(
                chunk,
                "KeyNotFoundException",
                line,
            );
            vybe_compiler::primitives::errors::emit_throw(chunk, line);
            chunk.emit_end(line);
        }
        "dotnet.dict_get_value_or_default" => {
            // Stack: [map, key] or [map, key, default]
            // map.has(key) ? map.get(key) : default
            let chunk = &mut chunks[current];
            let has = chunk.add_import("ecma:map", "has");
            let get = chunk.add_import("ecma:map", "get");
            if argc >= 3 {
                // Explicit default: [map, key, default]
                let default_slot = chunk.alloc_scratch(1);
                let key_slot = chunk.alloc_scratch(1);
                let map_slot = chunk.alloc_scratch(1);
                chunk.emit_op_u16(Op::LOCAL_SET, default_slot, line);
                chunk.emit_op_u16(Op::LOCAL_SET, key_slot, line);
                chunk.emit_op_u16(Op::LOCAL_SET, map_slot, line);
                chunk.emit_op_u16(Op::LOCAL_GET, map_slot, line);
                chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
                chunk.emit_call(has, 2, line);
                chunk.emit_if_value(line);
                chunk.emit_op_u16(Op::LOCAL_GET, map_slot, line);
                chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
                chunk.emit_call(get, 2, line);
                chunk.emit_else(line);
                chunk.emit_op_u16(Op::LOCAL_GET, default_slot, line);
                chunk.emit_end(line);
            } else {
                // No explicit default: [map, key] → default is 0
                let key_slot = chunk.alloc_scratch(1);
                let map_slot = chunk.alloc_scratch(1);
                chunk.emit_op_u16(Op::LOCAL_SET, key_slot, line);
                chunk.emit_op_u16(Op::LOCAL_SET, map_slot, line);
                chunk.emit_op_u16(Op::LOCAL_GET, map_slot, line);
                chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
                chunk.emit_call(has, 2, line);
                chunk.emit_if_value(line);
                chunk.emit_op_u16(Op::LOCAL_GET, map_slot, line);
                chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
                chunk.emit_call(get, 2, line);
                chunk.emit_else(line);
                chunk.emit_f64_const(0.0, line);
                chunk.emit_end(line);
            }
        }
        "dotnet.dict_try_get_value" => {
            // TryGetValue(key, out value) → has(key) ? (value=get(key), true) : (value=default, false)
            // Simplified: returns get(key) or null, caller checks
            let chunk = &mut chunks[current];
            let out_slot = chunk.alloc_scratch(1);
            let key_slot = chunk.alloc_scratch(1);
            let map_slot = chunk.alloc_scratch(1);
            chunk.emit_op_u16(Op::LOCAL_SET, out_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, key_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, map_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, map_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
            let has = chunk.add_import("ecma:map", "has");
            chunk.emit_call(has, 2, line);
            chunk.emit_if_value(line);
            chunk.emit_op_u16(Op::LOCAL_GET, map_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, key_slot, line);
            let get = chunk.add_import("ecma:map", "get");
            chunk.emit_call(get, 2, line);
            chunk.emit_else(line);
            chunk.emit_f64_const(0.0, line);
            chunk.emit_end(line);
        }
        "dotnet.concurrent_queue_try_dequeue" => {
            chunks[current].emit_op(Op::DROP, line);
            vybe_compiler::primitives::collections::emit_shift(chunks, current, line);
        }
        "dotnet.concurrent_stack_try_pop" => {
            chunks[current].emit_op(Op::DROP, line);
            vybe_compiler::primitives::collections::emit_pop(chunks, current, line);
        }
        "dotnet.concurrent_queue_try_peek" => {
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_i32_const(0, line);
            vybe_compiler::primitives::collections::emit_get(chunks, current, line);
        }
        "dotnet.concurrent_stack_try_peek" => {
            chunks[current].emit_op(Op::DROP, line);
            let recv = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
            vybe_compiler::primitives::collections::emit_len(chunks, current, line);
            chunks[current].emit_i32_const(1, line);
            chunks[current].emit_op(Op::I32_SUB, line);
            vybe_compiler::primitives::collections::emit_get(chunks, current, line);
        }
        "dotnet.choose" => emit_choose(&mut chunks[current], argc, line),
        "dotnet.string_compare" => {
            crate::emitter::core::string_adapter::emit_string_compare(chunks, current, argc, line)
        }
        "dotnet.string_equals" => {
            crate::emitter::core::string_adapter::emit_string_equals(chunks, current, argc, line)
        }
        "dotnet.string_contains" => {
            crate::emitter::core::string_adapter::emit_string_contains(chunks, current, argc, line)
        }
        "dotnet.string_starts_with" => {
            crate::emitter::core::string_adapter::emit_string_starts_with(
                chunks, current, argc, line,
            )
        }
        "dotnet.string_ends_with" => {
            crate::emitter::core::string_adapter::emit_string_ends_with(chunks, current, argc, line)
        }
        "dotnet.string_index_of" => {
            crate::emitter::core::string_adapter::emit_string_index_of(chunks, current, argc, line)
        }
        "dotnet.string_last_index_of" => {
            crate::emitter::core::string_adapter::emit_string_last_index_of(
                chunks, current, argc, line,
            )
        }
        "dotnet.string_index_of_any" => {
            crate::emitter::core::string_adapter::emit_string_index_of_any(
                chunks, current, argc, line,
            )
        }
        "dotnet.string_last_index_of_any" => {
            crate::emitter::core::string_adapter::emit_string_last_index_of_any(
                chunks, current, argc, line,
            )
        }
        "dotnet.string_substring" => {
            crate::emitter::core::string_adapter::emit_string_substring(chunks, current, argc, line)
        }
        "dotnet.string_char_at_checked" => {
            crate::emitter::core::string_adapter::emit_string_char_at_checked(chunks, current, line)
        }
        "dotnet.string_pad_left" => {
            crate::emitter::core::string_adapter::emit_string_pad_left(chunks, current, argc, line)
        }
        "dotnet.string_pad_right" => {
            crate::emitter::core::string_adapter::emit_string_pad_right(chunks, current, argc, line)
        }
        "dotnet.string_replace" => {
            crate::emitter::core::string_adapter::emit_string_replace(chunks, current, argc, line)
        }
        "dotnet.string_split" => {
            crate::emitter::core::string_adapter::emit_string_split(chunks, current, argc, line)
        }
        "dotnet.string_concat" => {
            crate::emitter::core::string_adapter::emit_string_concat(chunks, current, argc, line)
        }
        "dotnet.vb_strings_left" => {
            crate::emitter::core::string_adapter::emit_vb_strings_left(chunks, current, argc, line)
        }
        "dotnet.vb_strings_right" => {
            crate::emitter::core::string_adapter::emit_vb_strings_right(chunks, current, argc, line)
        }
        "dotnet.vb_strings_mid" => {
            crate::emitter::core::string_adapter::emit_vb_strings_mid(chunks, current, argc, line)
        }
        "dotnet.string_trim_chars" => crate::emitter::core::string_adapter::emit_string_trim_chars(
            chunks, current, argc, line,
        ),
        "dotnet.string_trim_start_chars" => {
            crate::emitter::core::string_adapter::emit_string_trim_start_chars(
                chunks, current, argc, line,
            )
        }
        "dotnet.string_trim_end_chars" => {
            crate::emitter::core::string_adapter::emit_string_trim_end_chars(
                chunks, current, argc, line,
            )
        }

        // ── System.Math — shared .NET BCL math surface ──────────────
        // WASM opcodes (zero overhead)
        "dotnet.system.math.abs" => chunks[current].emit_op(Op::F64_ABS, line),
        "dotnet.system.math.floor" => chunks[current].emit_op(Op::F64_FLOOR, line),
        "dotnet.system.math.ceiling" | "dotnet.system.math.ceil" => {
            chunks[current].emit_op(Op::F64_CEIL, line)
        }
        "dotnet.system.math.sqrt" => chunks[current].emit_op(Op::F64_SQRT, line),
        "dotnet.system.math.truncate" | "dotnet.system.math.trunc" => {
            chunks[current].emit_op(Op::F64_TRUNC, line)
        }
        "dotnet.system.math.round" => {
            if argc <= 1 {
                chunks[current].emit_op(Op::F64_NEAREST, line);
            } else {
                // Round(value, digits): nearest(value * 10^digits) / 10^digits
                let chunk = &mut chunks[current];
                let digits_slot = chunk.alloc_scratch(1);
                let val_slot = chunk.alloc_scratch(1);
                let factor_slot = chunk.alloc_scratch(1);
                chunk.emit_op_u16(Op::LOCAL_SET, digits_slot, line);
                chunk.emit_op_u16(Op::LOCAL_SET, val_slot, line);
                chunk.emit_f64_const(10.0, line);
                chunk.emit_op_u16(Op::LOCAL_GET, digits_slot, line);
                let pow = chunk.add_import("ecma:math", "pow");
                chunk.emit_call(pow, 2, line);
                chunk.emit_op_u16(Op::LOCAL_SET, factor_slot, line);
                chunk.emit_op_u16(Op::LOCAL_GET, val_slot, line);
                chunk.emit_op_u16(Op::LOCAL_GET, factor_slot, line);
                chunk.emit_op(Op::F64_MUL, line);
                chunk.emit_op(Op::F64_NEAREST, line);
                chunk.emit_op_u16(Op::LOCAL_GET, factor_slot, line);
                chunk.emit_op(Op::F64_DIV, line);
            }
        }
        "dotnet.system.math.min" => chunks[current].emit_op(Op::F64_MIN, line),
        "dotnet.system.math.max" => chunks[current].emit_op(Op::F64_MAX, line),
        // Host calls (ecma:math)
        "dotnet.system.math.pow" => {
            let idx = chunks[current].add_import("ecma:math", "pow");
            chunks[current].emit_call(idx, 2, line);
        }
        "dotnet.system.math.sin" => {
            let idx = chunks[current].add_import("ecma:math", "sin");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.cos" => {
            let idx = chunks[current].add_import("ecma:math", "cos");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.tan" => {
            let idx = chunks[current].add_import("ecma:math", "tan");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.asin" => {
            let idx = chunks[current].add_import("ecma:math", "asin");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.acos" => {
            let idx = chunks[current].add_import("ecma:math", "acos");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.atan" => {
            let idx = chunks[current].add_import("ecma:math", "atan");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.atan2" => {
            let idx = chunks[current].add_import("ecma:math", "atan2");
            chunks[current].emit_call(idx, 2, line);
        }
        "dotnet.system.math.log" => {
            let idx = chunks[current].add_import("ecma:math", "log");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.log10" => {
            let idx = chunks[current].add_import("ecma:math", "log10");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.log2" => {
            let idx = chunks[current].add_import("ecma:math", "log2");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.exp" => {
            let idx = chunks[current].add_import("ecma:math", "exp");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.sinh" => {
            let idx = chunks[current].add_import("ecma:math", "sinh");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.cosh" => {
            let idx = chunks[current].add_import("ecma:math", "cosh");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.tanh" => {
            let idx = chunks[current].add_import("ecma:math", "tanh");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.sign" => {
            let idx = chunks[current].add_import("ecma:math", "sign");
            chunks[current].emit_call(idx, 1, line);
        }
        "dotnet.system.math.clamp" => {
            vybe_compiler::primitives::math::emit_clamp(&mut chunks[current], line)
        }
        _ => return false,
    }
    true
}
