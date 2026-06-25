//! Node-compatible socket helpers hosted in the shared dotnet adapter layer.
//!
//! These are not `.NET` API shims. They live here because this crate's
//! shared network lowering already centers on the dotnet/core adapter
//! surface, and other frontends can reuse the same bytecode helpers.
//!
//! The adapters keep Node-shaped JS call signatures while lowering to the
//! real `wasi:sockets/*` surface. No `dotnet:*` host module is involved.

use crate::emitter::instructions::core_wasm;
use std::sync::Arc;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

/// `net.createConnection(port[, host])` / `net.connect(port[, host])`
/// normalize Node arg order to the shared socket adapter's
/// `(host, port)` shape.
pub fn emit_net_create_connection(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    match argc {
        0 => {
            push_const(chunk, Value::String(Arc::from("127.0.0.1")), line);
            core_wasm::i32_const(chunk, line, 0);
        }
        1 => {
            let port_slot = chunk.local_count;
            chunk.local_count = port_slot + 1;
            chunk.emit_op_u16(Op::LOCAL_SET, port_slot, line);
            push_const(chunk, Value::String(Arc::from("127.0.0.1")), line);
            chunk.emit_op_u16(Op::LOCAL_GET, port_slot, line);
        }
        _ => {
            let host_slot = chunk.local_count;
            let port_slot = host_slot + 1;
            chunk.local_count = port_slot + 1;
            chunk.emit_op_u16(Op::LOCAL_SET, host_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, port_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, host_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, port_slot, line);
        }
    }

    super::sockets_adapter::emit_tcp_client_new(chunks, current, line);
}

/// `net.createServer([listener])` materializes a listening socket on an
/// ephemeral port; the optional callback is ignored at compile time.
pub fn emit_net_create_server(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc > 0 {
        let listener_slot = chunk.local_count;
        chunk.local_count = listener_slot + 1;
        chunk.emit_op_u16(Op::LOCAL_SET, listener_slot, line);
    }
    core_wasm::i32_const(chunk, line, 0);
    super::sockets_adapter::emit_tcp_listener_new(chunks, current, line);
}

/// `dgram.createSocket(type)` currently lowers to a UDP socket bound on
/// an ephemeral port. The `type` argument is accepted for Node source
/// compatibility; the compiler uses the shared UDP adapter path.
pub fn emit_dgram_create_socket(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc > 0 {
        let kind_slot = chunk.local_count;
        chunk.local_count = kind_slot + 1;
        chunk.emit_op_u16(Op::LOCAL_SET, kind_slot, line);
    }
    core_wasm::i32_const(chunk, line, 0);
    super::sockets_adapter::emit_udp_client_new(chunks, current, line);
}
