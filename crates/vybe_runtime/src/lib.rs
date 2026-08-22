pub mod bigint;
pub mod capabilities;
pub mod chunk;
pub mod js_builtins; // wasm:js-* CG proposals (dispatcher)
pub mod js_primitive_builtins; // wasm:js-{number,boolean,undefined,symbol,bigint}
pub mod js_string_builtins; // wasm:js-string (merged js-string-builtins)
pub mod opcode;
pub mod scheduler;
pub mod type_recorder;
pub mod value;
pub mod vm;
// `impl VM` partials — extracted from vm.rs for readability. Each file is
// its own `impl VM { ... }` block operating on the same struct defined in
// vm.rs. Private to the crate; external consumers keep using `VM::*`.
pub(crate) mod calls;
pub mod canon_copy;
pub mod canon_flat;
pub mod canon_flat_values;
pub mod canon_layout;
pub mod canon_value;
pub mod cm_task;
pub mod debug;
pub mod debugger;
pub(crate) mod dispatch;
pub mod error;
pub mod event_loop;
pub mod fiber;
pub mod handle_table;
pub mod heap;
pub(crate) mod jspi;
pub mod module_record;
pub mod resources;
pub mod shared_memory;
pub(crate) mod simd;
pub(crate) mod threads;
pub mod typedef;
pub(crate) mod upvalues;
pub mod waitable;

pub use bigint::{BigIntRef, BigIntVal};
pub use chunk::{Chunk, Import};
pub use debugger::{DebugCommand, DebugEvent, DebugRequest, DebugResponse, Debugger};
pub use error::VMError;
pub use event_loop::EventLoop;
pub use module_record::{ExportEntry, ModuleKind, ModuleRecord, ModuleRequest, ModuleStatus};
pub use opcode::Op;
pub use typedef::{FieldDef, Method, ResourceTable, TypeDef, TypeRegistry};
pub use value::Properties;
pub use value::Value;
pub use vm::{HostContext, HostFn, ImportTarget, VM, VmSnapshot};
pub mod component;
pub mod component_model;
pub mod project;

// Plugin SDK + registry (folded in from the former `vybe_plugin` crate).
// `vybe_runtime` is the single registry every plugin registers into: the
// `Plugin` trait + `Framework` (the registration surface), the one init loop,
// the process-global language/hook registry, and the language `profile`.
// Language-agnostic class IR lives in `vybe_ast::class_normalize`.
pub mod framework;
pub mod namespaces;
pub mod profile;
pub mod registry;
pub use component::{
    BinaryLoader, Component, ExportImpl, FuncSig, ImportPolicy, Interface, Language, LinkResult,
    Linker, ModuleExport, ModuleResolver, ResolvedModule, ValType, register_binary_loader,
};
pub use framework::{
    Framework, Plugin, PluginEntry, finalize_plugins, finalize_registered_plugins, init_all,
    init_all_on_vm, init_all_on_vm_with_caps, init_all_registered, init_plugins, init_registered,
    init_registered_plugins, plugins,
};
pub use inventory;
pub use project::ProjectConfig;

// ── Declared host signatures ────────────────────────────────────────────────
//
// Process-global, because the CONSUMER is the compiler and the compiler never
// holds a VM. The namespace tree is global for exactly the same reason, and
// `mount_host_exports` already reads the capability context that way.
//
// Sparse ON PURPOSE. A function with no entry is UNDECLARED, which is the
// honest state of every registration that has not migrated to `HostFnDecl`.
// Nothing infers a signature from absence — an undeclared call is left alone,
// not assumed to take zero arguments.

static HOST_SIGNATURES: std::sync::OnceLock<
    std::sync::RwLock<
        std::collections::HashMap<
            (String, String),
            (FuncSig, Option<crate::vm::ResourceBinding>),
        >,
    >,
> = std::sync::OnceLock::new();

fn host_signatures()
-> &'static std::sync::RwLock<
    std::collections::HashMap<(String, String), (FuncSig, Option<crate::vm::ResourceBinding>)>,
> {
    HOST_SIGNATURES.get_or_init(Default::default)
}

/// Record what a host function declares about itself. Written by
/// `VM::register_host`; read by the compiler.
pub fn declare_host_signature(
    module: &str,
    name: &str,
    sig: FuncSig,
    resource: Option<crate::vm::ResourceBinding>,
) {
    if let Ok(mut map) = host_signatures().write() {
        map.insert((module.to_string(), name.to_string()), (sig, resource));
    }
}

/// The declared parameter count, or `None` when the function never declared
/// one. `None` means UNKNOWN — never zero.
pub fn declared_host_arity(module: &str, name: &str) -> Option<u8> {
    let map = host_signatures().read().ok()?;
    let (sig, _) = map.get(&(module.to_string(), name.to_string()))?;
    u8::try_from(sig.params.len()).ok()
}

/// The declared signature and resource binding, when there is one.
pub fn declared_host_signature(
    module: &str,
    name: &str,
) -> Option<(FuncSig, Option<crate::vm::ResourceBinding>)> {
    let map = host_signatures().read().ok()?;
    map.get(&(module.to_string(), name.to_string())).cloned()
}

/// The declared arity, when it CONTRADICTS `argc`. `None` means "no
/// complaint" — either undeclared, or declared and matching.
///
/// Extracted as a predicate so the check is testable directly: an emit site
/// that prints has no return value to assert on, and "no warning appeared" is
/// equally consistent with "the check passed" and "the check never ran".
pub fn host_arity_mismatch(module: &str, name: &str, argc: u8) -> Option<u8> {
    let (sig, resource) = declared_host_signature(module, name)?;
    let declared = u8::try_from(sig.params.len()).ok()?;
    if declared == argc {
        return None;
    }

    // A METHOD on a resource may be called without its handle, because the
    // RECEIVER is the handle.
    //
    // `web:dom.appendChild` declares `(borrow<document>, node, node)` and the
    // emitters that construct controls pass all three. But the spec spelling is
    // `parent.appendChild(child)` — two arguments — and the document is not part
    // of the Web API at all: it is derived from the element, which is what
    // `ownerDocument` means (DOM §4.4). Both calls reach the same function and
    // both are correct, so one fixed number cannot describe the call site.
    //
    // Only `declared - 1`, and only for a method that borrows self: any other
    // count is still a real mismatch, and the case this check exists to catch —
    // the last argument forgotten and a null arriving in its place — is
    // untouched, because that call is short by one on a function with NO
    // receiver to make up the difference.
    let borrows_self = resource
        .as_ref()
        .is_some_and(|binding| binding.borrows_self);
    if borrows_self && argc + 1 == declared {
        return None;
    }

    Some(declared)
}

#[cfg(test)]
mod host_signature_tests {
    use super::*;

    #[test]
    fn an_undeclared_function_is_never_reported() {
        // Unknown is not zero. Every registration that has not migrated to
        // `HostFnDecl` must be left alone, whatever it is called with.
        assert_eq!(host_arity_mismatch("web:dom", "neverDeclared", 7), None);
    }

    #[test]
    fn a_declared_function_reports_only_a_real_mismatch() {
        declare_host_signature(
            "test:iface",
            "appendChild",
            FuncSig {
                name: "append-child".into(),
                params: vec![
                    ValType::Borrow("document".into()),
                    ValType::Borrow("node".into()),
                    ValType::Borrow("node".into()),
                ],
                results: vec![],
            },
            None,
        );
        assert_eq!(host_arity_mismatch("test:iface", "appendChild", 3), None);
        // The bug this exists to catch: the child forgotten, a null passed in
        // its place, and the failure surfacing somewhere else entirely.
        assert_eq!(host_arity_mismatch("test:iface", "appendChild", 2), Some(3));
    }

    #[test]
    fn a_resource_method_may_be_called_without_its_handle() {
        // `parent.appendChild(child)` is the spec spelling and passes TWO
        // arguments; the document is derived from the element. The positional
        // emitters pass all three. Both are correct, so the declared count
        // cannot be the only accepted one.
        declare_host_signature(
            "test:method",
            "appendChild",
            FuncSig {
                name: "append-child".into(),
                params: vec![
                    ValType::Borrow("document".into()),
                    ValType::Borrow("node".into()),
                    ValType::Borrow("node".into()),
                ],
                results: vec![],
            },
            Some(crate::vm::ResourceBinding {
                resource: "document".into(),
                kind: crate::vm::ResourceMemberKind::Method,
                borrows_self: true,
            }),
        );
        assert_eq!(host_arity_mismatch("test:method", "appendChild", 3), None);
        assert_eq!(host_arity_mismatch("test:method", "appendChild", 2), None);
        // Short by TWO is still short: the receiver can only supply one handle,
        // so this is the forgotten-argument bug the check exists for.
        assert_eq!(host_arity_mismatch("test:method", "appendChild", 1), Some(3));
        // And too many is always wrong.
        assert_eq!(host_arity_mismatch("test:method", "appendChild", 4), Some(3));
    }

    #[test]
    fn a_declaration_carries_its_resource_binding() {
        declare_host_signature(
            "test:iface",
            "setTextContent",
            FuncSig {
                name: "set-text-content".into(),
                params: vec![ValType::Borrow("node".into()), ValType::String],
                results: vec![],
            },
            Some(crate::vm::ResourceBinding {
                resource: "node".into(),
                kind: crate::vm::ResourceMemberKind::Method,
                borrows_self: true,
            }),
        );
        let (sig, resource) = declared_host_signature("test:iface", "setTextContent").unwrap();
        assert_eq!(sig.params.len(), 2);
        let binding = resource.expect("declared as a method on a resource");
        assert_eq!(binding.resource, "node");
        assert!(binding.borrows_self, "a DOM op borrows, it does not consume");
    }
}
