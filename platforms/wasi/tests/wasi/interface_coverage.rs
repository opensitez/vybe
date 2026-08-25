//! Every function the WASI interfaces declare, asserted registered.
//!
//! WASI 0.3.1 is six packages: `wasi:random`, `wasi:clocks`, `wasi:sockets`,
//! `wasi:filesystem`, `wasi:cli` and `wasi:http`
//! (`proposals/WASI/specifications/wasi-0.3.1/Overview.md`). The tables below
//! are transcribed from the WIT in `proposals/WASI/proposals/<package>/wit/`,
//! one entry per declared function, so a missing interface fails here by name
//! instead of surfacing later as `Unresolved import` in whichever language
//! happened to reach for it.
//!
//! Only IMPORTS appear. A world's EXPORTS are what a guest implements and a
//! host calls, so registering them as host functions would be backwards:
//!
//!   * `wasi:cli/run.run` — exported by the `command` world
//!     (`cli/wit/command.wit`), and imported by nothing. The clear case.
//!   * `wasi:http/handler.handle` is NOT that case, though it reads like it:
//!     the `service` world EXPORTS it, but `middleware` IMPORTS it as well
//!     (`world middleware { include service; import handler; }`), and the spec
//!     notes a `client.send` import may be linked directly to a
//!     `handler.handle` export. A host satisfying `handler` for a middleware
//!     component is therefore in-spec, so it is listed as an import here.
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
    // `handler` is the `middleware` world's import — see the note at the top of
    // this file. `client.send` and `handler.handle` have deliberately identical
    // signatures, which is why one implementation answers both.
    ("wasi:http/handler", &["handle"]),
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

// ── The converse: nothing registered that the spec does not declare ─────────
//
// Everything above asserts spec ⊆ registered. That direction alone is what let
// `fs.rs` register thirty invented verbs — `readFile`, `pathGetTempPath`,
// `openFile` — inside the REAL `wasi:filesystem` namespace and stay invisible:
// no declared function was missing, so every test above stayed green while the
// namespace said `wasi:` about a surface that is not WASI at all.
//
// A guest compiled against the actual WIT cannot call those verbs, and a
// conforming runtime cannot satisfy them, so the `wasi:` prefix is a claim the
// implementation does not keep. Asserting registered ⊆ spec is what turns that
// from a matter of opinion into a name a test can print.

/// The six packages WASI 0.3.1 defines, each paired with its table above.
///
/// Read from the vendored WIT, not from a file header: every `package` decl
/// under `proposals/WASI/proposals/*/wit/` is `@0.3.1`, and none is `@0.3.0`.
/// The headers in `platforms/wasi/` variously claim 0.2.8 and 0.2.12 and are
/// simply stale.
const PACKAGES: &[(&str, &[(&str, &[&str])])] = &[
    ("wasi:random", RANDOM),
    ("wasi:clocks", CLOCKS),
    ("wasi:cli", CLI),
    ("wasi:filesystem", FILESYSTEM),
    ("wasi:http", HTTP),
    ("wasi:sockets", SOCKETS),
];

/// `wasi:*` packages that are NOT part of 0.3.1 and are not violations.
///
/// `wasi:crypto`, `wasi:sql` and `wasi:logging` are each at their own phase
/// with no WIT vendored in this tree, so there is nothing here to check them
/// against. Listing them is the honest position: they are out of scope for
/// 0.3.1, rather than either silently tolerated or wrongly failed.
///
/// `wasi:tls` is different and worth the distinction: its WIT IS vendored
/// (`proposals/wasi-tls/wit-0.3.0-draft/`) and its surface IS checked, by
/// `tls.rs`, against the five functions that draft declares. It appears here
/// only because it is not one of the six — not because nothing verifies it.
const SEPARATE_PROPOSALS: &[&str] =
    &["wasi:crypto", "wasi:sql", "wasi:logging", "wasi:tls"];

/// Packages 0.3.1 DELETED that this tree still registers.
///
/// `wasi:io` is gone in 0.3.1 — streams and futures became component-model
/// built-ins, so `stream<u8>` is a canonical type and needs no interface. It is
/// still registered in two places: `io.rs` (23 functions) and
/// `sockets.rs::register_wasi_io` (a further 17). Retiring it is a rewrite of
/// the sockets stream surface, not a deletion, because 0.3 sockets hand back a
/// `stream<u8>` directly where 0.2 handed back an `input-stream` resource.
///
/// It is named here rather than omitted so the debt is structural: DELETING
/// EMPTY, and that is the point.
///
/// `wasi:io` was the only entry, and it is now gone from the tree outright:
/// `platforms/wasi/src/io.rs` is DELETED, along with `register_wasi_io`, the
/// 0.2 `register_wasi_sockets`, `register_wasi_sockets_method_forms` and
/// `register_io_streams`. Deleting the entry and watching this file stay green
/// was the stated definition of done.
///
/// Keep the list rather than the concept: the next deleted-but-still-live
/// package needs somewhere to be NAMED, and an exemption no test mentions is
/// indistinguishable from an oversight. What must not happen again is an entry
/// sitting here long enough that the gate reads GREEN while a whole
/// non-existent package answers calls.
const RETIRING: &[&str] = &[];

/// Every registered host function whose module starts `wasi:`.
pub(crate) fn registered_wasi_names_for_test() -> std::collections::BTreeSet<(String, String)> {
    registered_wasi_names().into_iter().collect()
}

fn registered_wasi_names() -> Vec<(String, String)> {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let mut names: Vec<(String, String)> = vm
        .host_registry
        .keys()
        .filter(|(module, _)| module.starts_with("wasi:"))
        .cloned()
        .collect();
    names.sort();
    names
}

/// The package part of `wasi:filesystem/types` — everything before the `/`.
fn package_of(module: &str) -> &str {
    module.split('/').next().unwrap_or(module)
}

/// No function may be registered under one of the six 0.3.1 packages unless
/// that package's WIT declares it.
///
/// This is the assertion that names `fs.rs`'s invented verbs. It also catches
/// 0.2 spellings kept alongside their 0.3 replacements rather than instead of
/// them — `wasi:http/outgoing-handler.handle` (0.3.1 says `client.send`), the
/// nine 0.2 HTTP resources the four 0.3.1 ones replaced, and the five 0.2
/// socket interfaces now collapsed into `wasi:sockets/types`.
///
/// `wasi:clocks` was the first package brought to zero here, which is what the
/// list looks like when a package is actually done.
#[test]
fn the_six_packages_register_only_what_they_declare() {
    let mut undeclared: Vec<String> = Vec::new();
    for (module, name) in registered_wasi_names() {
        let Some((_, surface)) = PACKAGES
            .iter()
            .find(|(package, _)| package_of(&module) == *package)
        else {
            continue; // a different package; the next test judges those
        };
        let declared = surface
            .iter()
            .any(|(iface, names)| *iface == module && names.contains(&name.as_str()));
        if !declared {
            undeclared.push(format!("{module} {name}"));
        }
    }
    assert!(
        undeclared.is_empty(),
        "{} functions are registered under a WASI 0.3.1 package that does not \
         declare them. A guest compiled against the WIT cannot call these, and \
         a conforming runtime cannot satisfy them:\n  {}",
        undeclared.len(),
        undeclared.join("\n  ")
    );
}

/// No `wasi:` package may be registered that is neither one of the six, nor a
/// separate proposal, nor named as retiring debt.
#[test]
fn the_wasi_namespace_holds_only_accounted_packages() {
    let known: Vec<&str> = PACKAGES
        .iter()
        .map(|(package, _)| *package)
        .chain(SEPARATE_PROPOSALS.iter().copied())
        .chain(RETIRING.iter().copied())
        .collect();
    let mut stray: Vec<String> = registered_wasi_names()
        .into_iter()
        .map(|(module, _)| package_of(&module).to_string())
        .filter(|package| !known.contains(&package.as_str()))
        .collect();
    stray.sort();
    stray.dedup();
    assert!(
        stray.is_empty(),
        "unaccounted `wasi:` packages registered: {}",
        stray.join(", ")
    );
}
