use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::VM;
use vybe_runtime::capabilities::Capabilities;

fn has_import(module: &str, name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(module.to_string(), name.to_string()))
}

#[test]
fn proposal_tls_types_surface_is_registered() {
    let expected = ["[method]error.to-debug-string"];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi:tls/types", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing wasi-tls types imports: {missing:?}"
    );
}

#[test]
fn proposal_tls_client_surface_is_registered() {
    let expected = [
        "[constructor]connector",
        "[method]connector.send",
        "[method]connector.receive",
        "[static]connector.connect",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi:tls/client", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing wasi-tls client imports: {missing:?}"
    );
}
