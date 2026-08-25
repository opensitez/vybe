mod cli;
mod clocks;
mod crypto;
mod crypto_vectors;
mod filesystem;
mod filesystem_names;
mod filesystem_paths;
mod filesystem_stream_matrix;
mod filesystem_symlink_matrix;
mod http;
mod http_incoming_server;
mod http_request_lifecycle;
mod http_request_matrix;
mod http_spec_0_3;
mod http_spec_0_3_behaviour;
mod http_status_matrix;
mod interface_coverage;
// `io`, `io_contracts`, `io_length_matrix` and `io_poll_matrix` are DELETED,
// not skipped. Between them they made 98 calls to `wasi:io/{streams,poll,error}`
// and 40 to the 0.2 socket interfaces (`instance-network`, `tcp-create-socket`,
// `tcp`) — and NOT ONE to anything WASI 0.3.1 declares. There was nothing to
// rewrite: a test whose every call names a deleted interface is not a test of
// this system.
//
// The replacements already exist and are green: `stream_drain` +
// `filesystem_stream_matrix` cover reading a `stream<u8>` through
// `canon stream.read`, which is what `input-stream.read` became, and
// `sockets` + `sockets_contracts` cover the 0.3.1 socket surface.
mod random;
mod stream_drain;
mod surface_from_wit;
mod sockets;
mod sockets_contracts;
mod tls;
