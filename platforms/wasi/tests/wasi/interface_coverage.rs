//! Every function the WASI interfaces declare, asserted registered.
//!
//! WASI 0.3.0 (`specifications/wasi-0.3.0/Overview.md`) is six packages:
//! `wasi:random`, `wasi:clocks`, `wasi:sockets`, `wasi:filesystem`,
//! `wasi:cli` and `wasi:http`. The tables below are transcribed from the WIT
//! in `proposals/<package>/wit/`, one entry per declared function, so a
//! missing interface fails here by name instead of surfacing later as
//! `Unresolved import` in whichever language happened to reach for it.
//!
//! Only IMPORTS appear. A world's EXPORTS are what a guest implements and a
//! host calls, so registering them as host functions would be backwards:
//!
//!   * `wasi:cli/run.run` — exported by the `command` world (`cli/wit/command.wit`).
//!   * `wasi:http/handler.handle` — exported by the `service` world
//!     (`http/wit/worlds.wit`). Imported only by `middleware`, which forwards
//!     to another component rather than to the host; `client.send` is the
//!     host-provided direction and IS listed.
//!
//! Names follow the canonical component-model spelling the VM registry uses:
//! `[method]resource.name`, `[static]resource.name`, `[constructor]resource`.

use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::VM;
use vybe_runtime::capabilities::Capabilities;

/// `proposals/random/wit/{random,insecure,insecure-seed}.wit`
const RANDOM: &[(&str, &[&str])] = &[
    (
        "wasi:random/random",
        &["get-random-bytes", "get-random-u64"],
    ),
    (
        "wasi:random/insecure",
        &["get-insecure-random-bytes", "get-insecure-random-u64"],
    ),
    ("wasi:random/insecure-seed", &["get-insecure-seed"]),
];

/// `proposals/clocks/wit/{monotonic-clock,system-clock,timezone}.wit`
///
/// 0.3 renamed `wall-clock` to `system-clock` and `resolution` to
/// `get-resolution`; `timezone` gained `iana-id` and `to-debug-string`.
const CLOCKS: &[(&str, &[&str])] = &[
    (
        "wasi:clocks/monotonic-clock",
        &["now", "get-resolution", "wait-until", "wait-for"],
    ),
    ("wasi:clocks/system-clock", &["now", "get-resolution"]),
    (
        "wasi:clocks/timezone",
        &["iana-id", "utc-offset", "to-debug-string"],
    ),
];

/// `proposals/cli/wit/{environment,exit,stdio,terminal}.wit`
const CLI: &[(&str, &[&str])] = &[
    (
        "wasi:cli/environment",
        &["get-environment", "get-arguments", "get-initial-cwd"],
    ),
    ("wasi:cli/exit", &["exit", "exit-with-code"]),
    ("wasi:cli/stdin", &["read-via-stream"]),
    ("wasi:cli/stdout", &["write-via-stream"]),
    ("wasi:cli/stderr", &["write-via-stream"]),
    ("wasi:cli/terminal-stdin", &["get-terminal-stdin"]),
    ("wasi:cli/terminal-stdout", &["get-terminal-stdout"]),
    ("wasi:cli/terminal-stderr", &["get-terminal-stderr"]),
];

/// `proposals/filesystem/wit/{types,preopens}.wit`
const FILESYSTEM: &[(&str, &[&str])] = &[
    ("wasi:filesystem/preopens", &["get-directories"]),
    (
        "wasi:filesystem/types",
        &[
            "[method]descriptor.advise",
            "[method]descriptor.append-via-stream",
            "[method]descriptor.create-directory-at",
            "[method]descriptor.get-flags",
            "[method]descriptor.get-type",
            "[method]descriptor.is-same-object",
            "[method]descriptor.link-at",
            "[method]descriptor.metadata-hash",
            "[method]descriptor.metadata-hash-at",
            "[method]descriptor.open-at",
            "[method]descriptor.read-directory",
            "[method]descriptor.read-via-stream",
            "[method]descriptor.readlink-at",
            "[method]descriptor.remove-directory-at",
            "[method]descriptor.rename-at",
            "[method]descriptor.set-size",
            "[method]descriptor.set-times",
            "[method]descriptor.set-times-at",
            "[method]descriptor.stat",
            "[method]descriptor.stat-at",
            "[method]descriptor.symlink-at",
            "[method]descriptor.sync",
            "[method]descriptor.sync-data",
            "[method]descriptor.unlink-file-at",
            "[method]descriptor.write-via-stream",
        ],
    ),
];

/// `proposals/http/wit/{types,worlds}.wit`
///
/// 0.3 collapses 0.2's `incoming-request`/`outgoing-request` and
/// `incoming-response`/`outgoing-response` into one `request` and one
/// `response`, drops `incoming-body`/`outgoing-body` in favour of
/// `consume-body` returning a `stream<u8>`, and replaces `fields.entries`
/// with `copy-all`.
const HTTP: &[(&str, &[&str])] = &[
    (
        "wasi:http/types",
        &[
            "[constructor]fields",
            "[static]fields.from-list",
            "[method]fields.get",
            "[method]fields.has",
            "[method]fields.set",
            "[method]fields.delete",
            "[method]fields.get-and-delete",
            "[method]fields.append",
            "[method]fields.copy-all",
            "[method]fields.clone",
            "[static]request.new",
            "[method]request.get-method",
            "[method]request.set-method",
            "[method]request.get-path-with-query",
            "[method]request.set-path-with-query",
            "[method]request.get-scheme",
            "[method]request.set-scheme",
            "[method]request.get-authority",
            "[method]request.set-authority",
            "[method]request.get-options",
            "[method]request.get-headers",
            "[static]request.consume-body",
            "[static]response.new",
            "[method]response.get-status-code",
            "[method]response.set-status-code",
            "[method]response.get-headers",
            "[static]response.consume-body",
            "[constructor]request-options",
            "[method]request-options.get-connect-timeout",
            "[method]request-options.set-connect-timeout",
            "[method]request-options.get-first-byte-timeout",
            "[method]request-options.set-first-byte-timeout",
            "[method]request-options.get-between-bytes-timeout",
            "[method]request-options.set-between-bytes-timeout",
            "[method]request-options.clone",
        ],
    ),
    ("wasi:http/client", &["send"]),
];

/// `proposals/sockets/wit/{types,ip-name-lookup}.wit`
///
/// 0.3 replaces 0.2's `tcp`, `udp`, `tcp-create-socket`, `udp-create-socket`
/// and `instance-network` interfaces with a single `types`, where a socket is
/// made by `tcp-socket.create` / `udp-socket.create` and I/O is `send` /
/// `receive` over component-model streams.
const SOCKETS: &[(&str, &[&str])] = &[
    ("wasi:sockets/ip-name-lookup", &["resolve-addresses"]),
    (
        "wasi:sockets/types",
        &[
            "[static]tcp-socket.create",
            "[method]tcp-socket.bind",
            "[method]tcp-socket.connect",
            "[method]tcp-socket.listen",
            "[method]tcp-socket.send",
            "[method]tcp-socket.receive",
            "[method]tcp-socket.get-local-address",
            "[method]tcp-socket.get-remote-address",
            "[method]tcp-socket.get-is-listening",
            "[method]tcp-socket.get-address-family",
            "[method]tcp-socket.set-listen-backlog-size",
            "[method]tcp-socket.get-keep-alive-enabled",
            "[method]tcp-socket.set-keep-alive-enabled",
            "[method]tcp-socket.get-keep-alive-idle-time",
            "[method]tcp-socket.set-keep-alive-idle-time",
            "[method]tcp-socket.get-keep-alive-interval",
            "[method]tcp-socket.set-keep-alive-interval",
            "[method]tcp-socket.get-keep-alive-count",
            "[method]tcp-socket.set-keep-alive-count",
            "[method]tcp-socket.get-hop-limit",
            "[method]tcp-socket.set-hop-limit",
            "[method]tcp-socket.get-receive-buffer-size",
            "[method]tcp-socket.set-receive-buffer-size",
            "[method]tcp-socket.get-send-buffer-size",
            "[method]tcp-socket.set-send-buffer-size",
            "[static]udp-socket.create",
            "[method]udp-socket.bind",
            "[method]udp-socket.connect",
            "[method]udp-socket.disconnect",
            "[method]udp-socket.send",
            "[method]udp-socket.receive",
            "[method]udp-socket.get-local-address",
            "[method]udp-socket.get-remote-address",
            "[method]udp-socket.get-address-family",
            "[method]udp-socket.get-unicast-hop-limit",
            "[method]udp-socket.set-unicast-hop-limit",
            "[method]udp-socket.get-receive-buffer-size",
            "[method]udp-socket.set-receive-buffer-size",
            "[method]udp-socket.get-send-buffer-size",
            "[method]udp-socket.set-send-buffer-size",
        ],
    ),
];

fn missing_from(surface: &[(&str, &[&str])]) -> Vec<String> {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let mut missing = Vec::new();
    for (module, names) in surface {
        for name in *names {
            if !vm
                .host_registry
                .contains_key(&((*module).to_string(), (*name).to_string()))
            {
                missing.push(format!("{module} {name}"));
            }
        }
    }
    missing
}

fn assert_complete(package: &str, surface: &[(&str, &[&str])]) {
    let missing = missing_from(surface);
    let total: usize = surface.iter().map(|(_, names)| names.len()).sum();
    assert!(
        missing.is_empty(),
        "{package}: {} of {total} declared functions are not registered:\n  {}",
        missing.len(),
        missing.join("\n  ")
    );
}

#[test]
fn random_interface_is_fully_registered() {
    assert_complete("wasi:random", RANDOM);
}

#[test]
fn clocks_interface_is_fully_registered() {
    assert_complete("wasi:clocks", CLOCKS);
}

#[test]
fn cli_interface_is_fully_registered() {
    assert_complete("wasi:cli", CLI);
}

#[test]
fn filesystem_interface_is_fully_registered() {
    assert_complete("wasi:filesystem", FILESYSTEM);
}

#[test]
fn http_interface_is_fully_registered() {
    assert_complete("wasi:http", HTTP);
}

#[test]
fn sockets_interface_is_fully_registered() {
    assert_complete("wasi:sockets", SOCKETS);
}
