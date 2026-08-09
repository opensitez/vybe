//! Conformance: every function in the WASI 0.3 HTTP WIT is registered.
//!
//! Source of truth: `proposals/wasi-http/wit-0.3.0-draft/{types,worlds}.wit`
//! (`wasi:http@0.3.0-rc-2025-09-16`). Resource funcs use the Component Model
//! `[method]<resource>.<name>` / `[static]<resource>.<name>` naming.

use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::VM;
use vybe_runtime::capabilities::Capabilities;

fn registered() -> VM {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm
}

fn assert_all(module: &str, names: &[&str]) {
    let vm = registered();
    let missing: Vec<&str> = names
        .iter()
        .copied()
        .filter(|name| {
            !vm.host_registry
                .contains_key(&(module.to_string(), name.to_string()))
        })
        .collect();
    assert!(missing.is_empty(), "{module} missing: {missing:?}");
}

#[test]
fn wasi_http_0_3_fields_surface_is_registered() {
    assert_all(
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
        ],
    );
}

#[test]
fn wasi_http_0_3_request_surface_is_registered() {
    assert_all(
        "wasi:http/types",
        &[
            "[static]request.new",
            "[static]request.consume-body",
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
        ],
    );
}

#[test]
fn wasi_http_0_3_response_surface_is_registered() {
    assert_all(
        "wasi:http/types",
        &[
            "[static]response.new",
            "[static]response.consume-body",
            "[method]response.get-status-code",
            "[method]response.set-status-code",
            "[method]response.get-headers",
        ],
    );
}

#[test]
fn wasi_http_0_3_request_options_surface_is_registered() {
    assert_all(
        "wasi:http/types",
        &[
            "[constructor]request-options",
            "[method]request-options.get-connect-timeout",
            "[method]request-options.set-connect-timeout",
            "[method]request-options.get-first-byte-timeout",
            "[method]request-options.set-first-byte-timeout",
            "[method]request-options.get-between-bytes-timeout",
            "[method]request-options.set-between-bytes-timeout",
            "[method]request-options.clone",
        ],
    );
}

#[test]
fn wasi_http_0_3_client_and_handler_interfaces_are_registered() {
    // `interface client { send: async func(request) -> result<response, error-code> }`
    // and the identically-shaped `interface handler { handle: ... }`.
    assert_all("wasi:http/client", &["send"]);
    assert_all("wasi:http/handler", &["handle"]);
}
