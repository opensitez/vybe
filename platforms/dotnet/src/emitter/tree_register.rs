//! `dotnet.*` namespace-tree registration (namespaceplan.md, dotnet phase).
//!
//! The dotnet platform contributes DATA — its component-model class
//! descriptors — to the shared namespace tree in `vybe_emitter::namespaces`.
//! Resolution LOGIC lives only in the common resolver: VB, C#, and every
//! other language resolve `dotnet.system.console.writeline` through the
//! same tree walk, instead of a platform-owned dotted-name cascade
//! (`resolver.rs`, which this registration supersedes and which dissolves
//! once VB/C# routing is fully migrated).

use std::collections::BTreeMap;
use std::sync::Once;

use vybe_bytecode::component_model::MethodBody;
use vybe_emitter::namespaces::{self, NamespaceNode, Subtree};

/// Register every component class descriptor as a `Type` node at
/// `<interface path>.<class name>` — statics become `CommonEmit`/host-fn
/// leaves (the same `MethodBody` targets `dispatch.rs` executes).
/// Idempotent; first call wins.
pub fn register_namespace_tree() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        for export in super::class_exports::dotnet_class_exports() {
            let (interface, class) = (export.interface, &export.class);
            let mut statics = Subtree::new();
            for m in &class.methods {
                if !m.is_static {
                    // Instance methods dispatch receiver-based (TypeRegistry
                    // vtables), never through the namespace tree.
                    continue;
                }
                let node = match &m.body {
                    MethodBody::Common(emit) => NamespaceNode::CommonEmit(emit.clone()),
                    MethodBody::HostCall(t) => namespaces::host_fn(&t.module, &t.name),
                    // Chunk-backed methods are per-compilation artifacts,
                    // not process-global surface.
                    MethodBody::UserChunk(_) => continue,
                };
                statics.insert(m.name.to_lowercase(), node);
            }
            // Host-backed property getters — the legacy static-property
            // lookup surface (`Stopwatch.Frequency`-style reads through a
            // class path) as tree leaves. Methods win on name collision.
            for p in &class.properties {
                if let Some(getter) = &p.getter {
                    statics
                        .entry(p.name.to_lowercase())
                        .or_insert_with(|| namespaces::host_fn(&getter.module, &getter.name));
                }
            }
            if interface.eq_ignore_ascii_case("dotnet.System")
                && class.name.eq_ignore_ascii_case("Console")
            {
                statics.insert("out".into(), console_stdout_writer_node());
                statics.insert("error".into(), console_stderr_writer_node());
            }
            if interface.eq_ignore_ascii_case("dotnet.System")
                && class.name.eq_ignore_ascii_case("Object")
            {
                statics.insert(
                    "equals".into(),
                    NamespaceNode::CommonEmit("dotnet.object_equals".into()),
                );
                statics.insert(
                    "referenceequals".into(),
                    NamespaceNode::CommonEmit("dotnet.object_reference_equals".into()),
                );
            }
            let ty = NamespaceNode::Type {
                ctor: None,
                statics,
                methods: BTreeMap::new(),
            };

            // "dotnet.System" + "Math" → dotnet.system.math
            let mut segments: Vec<String> =
                interface.split('.').map(|s| s.to_lowercase()).collect();
            segments.push(class.name.to_lowercase());

            let mut node = ty;
            while segments.len() > 1 {
                let key = segments.pop().expect("non-empty");
                let mut children = Subtree::new();
                children.insert(key, node);
                node = NamespaceNode::Namespace(children);
            }
            namespaces::register_namespace_tree(&segments.pop().expect("root"), node);
        }
    });
}

fn console_stdout_writer_node() -> NamespaceNode {
    let mut statics = Subtree::new();
    statics.insert(
        "write".into(),
        NamespaceNode::CommonEmit("dotnet.console_write".into()),
    );
    statics.insert(
        "writeline".into(),
        NamespaceNode::CommonEmit("dotnet.console_writeline".into()),
    );
    NamespaceNode::Namespace(statics)
}

fn console_stderr_writer_node() -> NamespaceNode {
    let mut statics = Subtree::new();
    statics.insert(
        "write".into(),
        NamespaceNode::CommonEmit("dotnet.console_error_write".into()),
    );
    statics.insert(
        "writeline".into(),
        NamespaceNode::CommonEmit("dotnet.console_error_writeline".into()),
    );
    NamespaceNode::Namespace(statics)
}

#[cfg(test)]
mod resolve_gap_tests {
    use vybe_emitter::namespaces::{ResolutionTarget, resolve_path};

    #[test]
    fn delegate_combine_resolves_via_tree() {
        super::register_namespace_tree();
        let r = resolve_path(&["dotnet", "system", "delegate", "combine"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => assert_eq!(name, "delegates.combine"),
            other => panic!("expected CommonEmit(delegates.combine), got {other:?}"),
        }
    }

    #[test]
    fn guid_parse_resolves_via_tree() {
        super::register_namespace_tree();
        let r = resolve_path(&["dotnet", "system", "guid", "parse"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => assert_eq!(name, "dotnet.guid_parse"),
            other => panic!("expected CommonEmit(dotnet.guid_parse), got {other:?}"),
        }
    }

    #[test]
    fn environment_set_resolves_via_tree() {
        super::register_namespace_tree();
        let r = resolve_path(&["dotnet", "system", "environment", "setenvironmentvariable"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => assert_eq!(name, "dotnet.environment_set"),
            other => panic!("expected CommonEmit(dotnet.environment_set), got {other:?}"),
        }
    }

    #[test]
    fn convert_and_parse_resolve_via_tree() {
        super::register_namespace_tree();

        let r = resolve_path(&["dotnet", "system", "convert", "tostring"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => assert_eq!(name, "dotnet.tostring"),
            other => panic!("expected CommonEmit(dotnet.tostring), got {other:?}"),
        }

        let r = resolve_path(&["dotnet", "system", "int32", "parse"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => assert_eq!(name, "dotnet.parse_int"),
            other => panic!("expected CommonEmit(dotnet.parse_int), got {other:?}"),
        }

        let r = resolve_path(&["dotnet", "system", "int", "tryparse"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => assert_eq!(name, "dotnet.try_parse_int"),
            other => panic!("expected CommonEmit(dotnet.try_parse_int), got {other:?}"),
        }
    }

    #[test]
    fn math_methods_resolve_via_tree() {
        super::register_namespace_tree();

        let r = resolve_path(&["dotnet", "system", "math", "sin"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => assert_eq!(name, "dotnet.system.math.sin"),
            other => panic!("expected CommonEmit(dotnet.system.math.sin), got {other:?}"),
        }

        let r = resolve_path(&["dotnet", "system", "math", "ceil"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => {
                assert_eq!(name, "dotnet.system.math.ceiling")
            }
            other => panic!("expected CommonEmit(dotnet.system.math.ceiling), got {other:?}"),
        }
    }

    #[test]
    fn console_and_datetime_resolve_via_tree() {
        super::register_namespace_tree();

        let r = resolve_path(&["dotnet", "system", "object", "referenceequals"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => {
                assert_eq!(name, "dotnet.object_reference_equals")
            }
            other => panic!("expected CommonEmit(dotnet.object_reference_equals), got {other:?}"),
        }

        let r = resolve_path(&["dotnet", "system", "console", "writeline"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => {
                assert_eq!(name, "dotnet.console_writeline")
            }
            other => panic!("expected CommonEmit(dotnet.console_writeline), got {other:?}"),
        }

        let r = resolve_path(&["dotnet", "system", "console", "error", "write"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => {
                assert_eq!(name, "dotnet.console_error_write")
            }
            other => panic!("expected CommonEmit(dotnet.console_error_write), got {other:?}"),
        }

        let r = resolve_path(&["dotnet", "system", "console", "out", "writeline"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => {
                assert_eq!(name, "dotnet.console_writeline")
            }
            other => panic!("expected CommonEmit(dotnet.console_writeline), got {other:?}"),
        }

        let r = resolve_path(&["dotnet", "system", "datetime", "now"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => assert_eq!(name, "dotnet.datetime_now"),
            other => panic!("expected CommonEmit(dotnet.datetime_now), got {other:?}"),
        }

        let r = resolve_path(&["dotnet", "system", "datetime", "parse"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => assert_eq!(name, "dotnet.datetime_parse"),
            other => panic!("expected CommonEmit(dotnet.datetime_parse), got {other:?}"),
        }
    }

    #[test]
    fn string_statics_resolve_via_tree() {
        super::register_namespace_tree();

        let r = resolve_path(&["dotnet", "system", "string", "format"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => assert_eq!(name, "dotnet.string_format"),
            other => panic!("expected CommonEmit(dotnet.string_format), got {other:?}"),
        }

        let r = resolve_path(&["dotnet", "system", "string", "join"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => {
                assert_eq!(name, "collections.join_sep_first")
            }
            other => panic!("expected CommonEmit(collections.join_sep_first), got {other:?}"),
        }

        let r = resolve_path(&["dotnet", "system", "string", "concat"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => assert_eq!(name, "str_concat"),
            other => panic!("expected CommonEmit(str_concat), got {other:?}"),
        }

        let r = resolve_path(&["dotnet", "system", "string", "isnullorempty"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => {
                assert_eq!(name, "dotnet.string_is_null_or_empty")
            }
            other => panic!("expected CommonEmit(dotnet.string_is_null_or_empty), got {other:?}"),
        }
    }

    #[test]
    fn system_text_encoding_resolves_via_tree() {
        super::register_namespace_tree();

        let r = resolve_path(&["dotnet", "system", "text", "encoding", "utf8"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => assert_eq!(name, "dotnet.encoding_utf8"),
            other => panic!("expected CommonEmit(dotnet.encoding_utf8), got {other:?}"),
        }

        let r = resolve_path(&["dotnet", "system", "text", "encoding", "getencoding"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => {
                assert_eq!(name, "dotnet.encoding_get_encoding")
            }
            other => panic!("expected CommonEmit(dotnet.encoding_get_encoding), got {other:?}"),
        }
    }

    #[test]
    fn io_threading_network_and_collections_resolve_via_tree() {
        super::register_namespace_tree();

        let r = resolve_path(&["dotnet", "system", "io", "file", "readalltext"]);
        match r {
            Some(ResolutionTarget::HostCall { module, func, .. }) => {
                assert_eq!(module, "wasi:filesystem");
                assert_eq!(func, "readFile");
            }
            other => panic!("expected HostCall(wasi:filesystem::readFile), got {other:?}"),
        }

        let r = resolve_path(&["dotnet", "system", "io", "directory", "getfiles"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => {
                assert_eq!(name, "dotnet.directory_get_files")
            }
            other => panic!("expected CommonEmit(dotnet.directory_get_files), got {other:?}"),
        }

        let r = resolve_path(&["dotnet", "system", "io", "path", "combine"]);
        match r {
            Some(ResolutionTarget::HostCall { module, func, .. }) => {
                assert_eq!(module, "wasi:filesystem");
                assert_eq!(func, "pathCombine");
            }
            other => panic!("expected HostCall(wasi:filesystem::pathCombine), got {other:?}"),
        }

        let r = resolve_path(&["dotnet", "system", "diagnostics", "stopwatch", "startnew"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => {
                assert_eq!(name, "dotnet.stopwatch_start_new")
            }
            other => panic!("expected CommonEmit(dotnet.stopwatch_start_new), got {other:?}"),
        }

        let r = resolve_path(&["dotnet", "system", "threading", "thread", "sleep"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => assert_eq!(name, "threading.sleep"),
            other => panic!("expected CommonEmit(threading.sleep), got {other:?}"),
        }

        let r = resolve_path(&["dotnet", "system", "net", "dns", "gethostname"]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => {
                assert_eq!(name, "dotnet.dns_get_host_name")
            }
            other => panic!("expected CommonEmit(dotnet.dns_get_host_name), got {other:?}"),
        }
    }

    #[test]
    fn winforms_application_resolves_via_tree() {
        super::register_namespace_tree();

        let r = resolve_path(&[
            "dotnet",
            "system",
            "windows",
            "forms",
            "application",
            "enablevisualstyles",
        ]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => {
                assert_eq!(name, "dotnet.winforms_noop");
            }
            other => panic!("expected CommonEmit(dotnet.winforms_noop), got {other:?}"),
        }

        let r = resolve_path(&[
            "dotnet",
            "system",
            "windows",
            "forms",
            "application",
            "setcompatibletextrenderingdefault",
        ]);
        match r {
            Some(ResolutionTarget::CommonEmit(name)) => {
                assert_eq!(name, "dotnet.winforms_noop");
            }
            other => panic!("expected CommonEmit(dotnet.winforms_noop), got {other:?}"),
        }
    }
}
