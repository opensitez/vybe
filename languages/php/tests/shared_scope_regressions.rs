// Exercise the production PHP/compiler-service dependency graph without the
// compiler crate's CLI dev-dependency bringing every language into this gate.
#[path = "../../../crates/vybe_compiler/tests/request_superglobal_bootstrap.rs"]
mod request_storage;
#[path = "../../../crates/vybe_compiler/tests/common_null_coalesce.rs"]
mod common_ast;
