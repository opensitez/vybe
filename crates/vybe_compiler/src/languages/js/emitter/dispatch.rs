//! Auto-extracted `js.*` dispatch (language-specific routing lives in the
//! language module; the common dispatcher delegates here).

use vybe_bytecode::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        "js.net_create_connection" => {
            crate::emitter::dotnet::core::node_socket_adapter::emit_net_create_connection(
                chunks, current, argc, line,
            )
        }
        "js.net_create_server" => {
            crate::emitter::dotnet::core::node_socket_adapter::emit_net_create_server(
                chunks, current, argc, line,
            )
        }
        "js.dgram_create_socket" => {
            crate::emitter::dotnet::core::node_socket_adapter::emit_dgram_create_socket(
                chunks, current, argc, line,
            )
        }

        // ── Threading ops ──
        // Real WASM threading opcodes (wasi-threads proposal):
        // thread_spawn, thread_join, memory atomic_*. NOT host calls — these
        // run unchanged on any standard WASM runtime that supports the
        // threads proposal.
        _ => return false,
    }
    true
}
