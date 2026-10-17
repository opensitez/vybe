// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml — see build.rs.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));
pub mod walker;

mod native_grammar {
    include!(concat!(env!("OUT_DIR"), "/native_parser.rs"));
}
pub(crate) use native_grammar::Rule;
pub(crate) struct WastParser;
impl WastParser {
    pub(crate) fn parse(
        rule: Rule,
        source: &str,
    ) -> Result<grameo::compat::Pairs<'_, Rule>, grameo::ParseError> {
        native_grammar::Parser.parse_pairs(rule, source)
    }
}

/// Parse WAST/WAT source into the common AST.
/// Both .wast (full script with assertions) and .wat (module only) are handled
/// by the same grammar — WAT is a strict subset of WAST.
pub fn parse(source: &str) -> Result<vybe_ast::Module, String> {
    walker::parse(source)
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}

/// Register this language with the shared plugin registry (dylib entry point).
pub fn register() {
    vybe_runtime::registry::register_language(vybe_runtime::registry::LanguageDef {
        name: "wast",
        parse,
        profile_source,
        emit_dispatch: None,
        normalize_class: None,
        register_tree: None,
        expand_source: None,
    });
}

/// This crate as a [`vybe_runtime::Plugin`] — its `init` registers the
/// language (and any forms) with the shared framework. Also the dylib entry point.
pub struct Plugin;
impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "wast"
    }
    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);

#[cfg(test)]
mod tests {
    #[test]
    fn value_type_export_emits_standard_component_binary() {
        let module = super::parse("(component (type $t u32) (export \"t\" (type $t)))")
            .expect("value type export parses");
        let binary = module.canon.binary.expect("ordered component definitions");
        let bytes =
            vybe_platform_wasm::write_component_binary(&binary).expect("type export validates");
        assert_eq!(
            vybe_platform_wasm::read_component_binary(&bytes).unwrap(),
            *binary
        );
    }

    #[test]
    fn composite_type_export_uses_normalized_component_types() {
        let module = super::parse(
            "(component (type $numbers (list u32)) \
             (type $item (record (field \"scores\" $numbers) (field \"name\" string))) \
             (export \"item\" (type $item)))",
        )
        .expect("composite component parses");
        let binary = module.canon.binary.expect("ordered composite definitions");
        let bytes = vybe_platform_wasm::write_component_binary(&binary)
            .expect("composite component validates");
        assert_eq!(
            vybe_platform_wasm::read_component_binary(&bytes).unwrap(),
            *binary
        );
    }

    #[test]
    fn type_export_alias_can_be_used_by_later_value_type() {
        let module = super::parse(
            "(component (type $value u32) (export $public \"value\" (type $value)) \
             (type $values (list $public)) (export \"values\" (type $values)))",
        )
        .expect("type export alias parses");
        let binary = module.canon.binary.expect("ordered alias definitions");
        let bytes =
            vybe_platform_wasm::write_component_binary(&binary).expect("alias type validates");
        assert_eq!(
            vybe_platform_wasm::read_component_binary(&bytes).unwrap(),
            *binary
        );
    }

    #[test]
    fn sync_and_async_function_types_emit_component_binary() {
        let module = super::parse(
            "(component (type $plain (func (param \"x\" u32) (result string))) \
             (type $async (func async (param \"x\" u32) (result u32))) \
             (export \"plain\" (type $plain)))",
        )
        .expect("function types parse");
        let binary = module
            .canon
            .binary
            .expect("ordered function type definitions");
        let bytes = vybe_platform_wasm::write_component_binary(&binary)
            .expect("function type binary validates");
        assert_eq!(
            vybe_platform_wasm::read_component_binary(&bytes).unwrap(),
            *binary
        );
    }

    #[test]
    fn resource_and_handle_types_emit_component_binary() {
        let module = super::parse(
            "(component (type $r (resource (rep i32))) \
             (type $own (own $r)) (type $borrow (borrow $r)))",
        )
        .expect("resource types parse");
        let binary = module.canon.binary.expect("ordered resource definitions");
        let bytes = vybe_platform_wasm::write_component_binary(&binary)
            .expect("resource component validates");
        assert_eq!(
            vybe_platform_wasm::read_component_binary(&bytes).unwrap(),
            *binary
        );
    }

    #[test]
    fn imported_function_lowering_emits_component_binary() {
        let module = super::parse(
            "(component (type $ft (func)) \
             (import \"f\" (func $f (type $ft))) \
             (core func $g (canon lower (func $f))) \
             (export \"run\" (func $f)))",
        )
        .expect("component import and lower parse");
        let binary = module.canon.binary.expect("ordered component definitions");
        let bytes = vybe_platform_wasm::write_component_binary(&binary)
            .expect("imported function component validates");
        assert_eq!(
            vybe_platform_wasm::read_component_binary(&bytes).unwrap(),
            *binary
        );
    }

    #[test]
    fn standalone_canon_lower_emits_component_binary() {
        let module = super::parse(
            "(component (type $ft (func)) \
             (import \"f\" (func $f (type $ft))) \
             (canon lower (func $f) (core func $g)))",
        )
        .expect("standalone canonical lower parses");
        let binary = module.canon.binary.expect("ordered canonical definition");
        let bytes = vybe_platform_wasm::write_component_binary(&binary)
            .expect("canonical lower component validates");
        assert_eq!(
            vybe_platform_wasm::read_component_binary(&bytes).unwrap(),
            *binary
        );
    }

    #[test]
    fn nested_component_function_import_links_and_emits_binary() {
        let module = super::parse(
            "(component
               (type $ft (func))
               (import \"f\" (func $f (type $ft)))
               (component $child
                 (type $inner_ft (func))
                 (import \"x\" (func $x (type $inner_ft)))
                 (export \"go\" (func $x)))
               (instance $i (instantiate $child (with \"x\" (func $f))))
               (alias export $i \"go\" (func $g))
               (export \"run\" (func $g)))",
        )
        .expect("nested function import should link to its supplied function");
        assert!(!module.canon.func_links.is_empty());
        let binary = module.canon.binary.expect("ordered component definitions");
        let bytes = vybe_platform_wasm::write_component_binary(&binary)
            .expect("nested component binary validates");
        assert_eq!(vybe_platform_wasm::read_component_binary(&bytes).unwrap(), *binary);
    }

    #[test]
    fn nested_component_missing_import_is_a_link_error() {
        let error = super::parse(
            "(component (component $child (type $ft (func)) \
             (import \"x\" (func $x (type $ft))) (export \"go\" (func $x))) \
             (instance $i (instantiate $child)))",
        )
        .expect_err("child import must be supplied");
        assert!(error.contains("missing function import `x`"), "{error}");
    }

    #[test]
    fn nested_component_import_invokes_supplied_lift() {
        let module = super::parse(
            "(component
               (type $ft (func (result u32)))
               (core module $m (func (export \"get\") (result i32) (i32.const 42)))
               (core instance $mi (instantiate $m))
               (alias core export $mi \"get\" (core func $get))
               (canon lift (core func $get) (func $parent (type $ft)))
               (component $child
                 (type $inner_ft (func (result u32)))
                 (import \"x\" (func $x (type $inner_ft)))
                 (export \"go\" (func $x)))
               (instance $i (instantiate $child (with \"x\" (func $parent))))
               (alias export $i \"go\" (func $g))
               (export \"run\" (func $g)))
             (assert_return (invoke \"run\") (i32.const 42))",
        )
        .expect("nested component source parses");
        let profile = vybe_compiler::profile::parse_profile(super::profile_source())
            .expect("WAST profile");
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile(&module)
            .expect("nested component compiles");
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
            .expect("nested component caller emits standard Wasm");
        let mut vm = vybe_runtime::VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        vm.run(chunks).expect("supplied function executes");
    }

    #[test]
    fn nested_component_import_resolves_inline_instance_export_path() {
        let module = super::parse(include_str!(
            "../../../tests/wast/wat_component/test_inline_instance_export_path.wast"
        ))
        .expect("inline instance export path parses");
        let profile = vybe_compiler::profile::parse_profile(super::profile_source())
            .expect("WAST profile");
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile(&module)
            .expect("inline alias compiles");
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
            .expect("inline alias caller emits standard Wasm");
        let mut vm = vybe_runtime::VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        vm.run(chunks).expect("inline instance export executes");
    }

    #[test]
    fn inline_instance_export_path_emits_standard_alias() {
        let module = super::parse(
            "(component
               (type $ft (func))
               (import \"f\" (func $f (type $ft)))
               (component $provider
                 (type $inner_ft (func))
                 (import \"x\" (func $x (type $inner_ft)))
                 (export \"produce\" (func $x)))
               (instance $p (instantiate $provider (with \"x\" (func $f))))
               (component $consumer
                 (type $inner_ft (func))
                 (import \"x\" (func $x (type $inner_ft)))
                 (export \"go\" (func $x)))
               (instance $c (instantiate $consumer
                 (with \"x\" (func $p \"produce\"))))
               (alias export $c \"go\" (func $run))
               (export \"run\" (func $run)))",
        )
        .expect("inline instance export path parses");
        let binary = module.canon.binary.as_ref().expect("ordered binary");
        let bytes = vybe_platform_wasm::write_component_binary(binary)
            .expect("inline alias binary validates");
        assert_eq!(vybe_platform_wasm::read_component_binary(&bytes).unwrap(), **binary);
    }

    #[test]
    fn nested_component_lower_calls_supplied_import() {
        let module = super::parse(include_str!(
            "../../../tests/wast/wat_component/test_nested_import_lower_call.wast"
        ))
        .expect("nested lower source parses");
        let profile = vybe_compiler::profile::parse_profile(super::profile_source())
            .expect("WAST profile");
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile(&module)
            .expect("nested lower compiles");
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&vybe_platform_wasm::write_wasm(&chunks))
            .expect("nested lower caller emits standard Wasm");
        let mut vm = vybe_runtime::VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        vm.run(chunks).expect("lowered imported function executes");
    }

    #[test]
    fn nested_component_import_rejects_mismatched_function_type() {
        let module = super::parse(
            "(component
               (type $text (func (result string)))
               (import \"supplied\" (func $f (type $text)))
               (component $child
                 (type $number (func (result u32)))
                 (import \"x\" (func $x (type $number))))
               (instance $i (instantiate $child (with \"x\" (func $f)))))",
        )
        .expect("mismatched import is syntactically valid");
        let profile = vybe_compiler::profile::parse_profile(super::profile_source())
            .expect("WAST profile");
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile(&module)
            .expect("link validation occurs in the VM");
        let error = vybe_runtime::VM::new()
            .run(chunks)
            .expect_err("incompatible function import must fail at instantiation");
        assert!(error.message.contains("type does not match"), "{error}");
    }

    #[test]
    fn empty_core_instance_emits_component_layer_binary() {
        let module =
            super::parse("(component (core module $m) (core instance $i (instantiate $m)))")
                .expect("component parses");
        assert!(module.canon.component_source);
        let binary = module.canon.binary.expect("ordered component definitions");
        let bytes = vybe_platform_wasm::write_component_binary(&binary)
            .expect("component binary validates");
        assert_eq!(&bytes[..8], &[0, 0x61, 0x73, 0x6d, 0x0d, 0, 1, 0]);
    }

    #[test]
    fn inline_empty_core_instance_emits_component_layer_binary() {
        let module = super::parse("(component (core instance $empty))")
            .expect("inline core instance parses");
        let binary = module.canon.binary.expect("ordered component definitions");
        assert!(matches!(
            binary.definitions.as_slice(),
            [vybe_compiler::ast::canon::ComponentBinaryDef::CoreInstanceFromExports(exports)]
                if exports.is_empty()
        ));
        let bytes = vybe_platform_wasm::write_component_binary(&binary)
            .expect("component binary validates");
        assert_eq!(
            vybe_platform_wasm::read_component_binary(&bytes).unwrap(),
            *binary
        );
    }

    #[test]
    fn core_instance_with_reference_round_trips_component_binary() {
        let module = super::parse(
            "(component (core module $m) (core instance $base) \
             (core instance $i (instantiate $m (with \"dep\" (instance $base)))))",
        )
        .expect("named core instance argument parses");
        let binary = module.canon.binary.expect("ordered component definitions");
        assert!(matches!(
            binary.definitions.last(),
            Some(vybe_compiler::ast::canon::ComponentBinaryDef::CoreInstance { module: 0, args })
                if args == &[("dep".to_string(), 0)]
        ));
        let bytes = vybe_platform_wasm::write_component_binary(&binary)
            .expect("component argument validates");
        assert_eq!(
            vybe_platform_wasm::read_component_binary(&bytes).unwrap(),
            *binary
        );
    }

    #[test]
    fn component_import_name_with_invalid_utf8_fails_closed() {
        let error = super::parse(
            r#"(component (core module $m) (core instance $base)
               (core instance (instantiate $m (with "\80" (instance $base)))))"#,
        )
        .unwrap_err();
        assert!(error.contains("not valid UTF-8"), "{error}");
    }

    #[test]
    fn component_type_labels_with_invalid_utf8_are_malformed() {
        for body in ["(record (field \"\\80\" s32))", "(flags \"\\80\")"] {
            let error = super::parse(&format!("(component (type {body}))")).unwrap_err();
            assert!(error.contains("not valid UTF-8"), "{body}: {error}");
        }
    }

    #[test]
    fn component_import_name_preserves_valid_escaped_utf8() {
        let module = super::parse(
            r#"(component (core module $m) (core instance $base)
               (core instance (instantiate $m (with "\c3\a9" (instance $base)))))"#,
        )
        .expect("valid UTF-8 import name parses");
        let binary = module.canon.binary.expect("ordered component definitions");
        assert!(matches!(
            binary.definitions.last(),
            Some(vybe_compiler::ast::canon::ComponentBinaryDef::CoreInstance { args, .. })
                if args == &[("\u{e9}".to_string(), 0)]
        ));
        let bytes = vybe_platform_wasm::write_component_binary(&binary)
            .expect("escaped UTF-8 import validates");
        assert_eq!(
            vybe_platform_wasm::read_component_binary(&bytes).unwrap(),
            *binary
        );
    }

    #[test]
    fn component_binary_metadata_survives_compilation() {
        let profile =
            vybe_compiler::profile::parse_profile(super::profile_source()).expect("WAST profile");
        let module =
            super::parse("(component (core module $m) (core instance $i (instantiate $m)))")
                .expect("supported component parses");
        let compiled = vybe_compiler::primitives::Compiler::with_profile(profile.clone())
            .compile_with_imports(&module)
            .expect("supported component compiles");
        assert!(compiled.component_source);
        assert_eq!(compiled.component_binary, module.canon.binary);
        assert!(std::sync::Arc::ptr_eq(
            compiled.component_binary.as_ref().unwrap(),
            module.canon.binary.as_ref().unwrap(),
        ));
        let bytes = vybe_platform_wasm::write_artifact(
            &compiled.chunks,
            compiled.component_source,
            compiled.component_binary.as_deref(),
            Default::default(),
        )
        .expect("supported source emits a component");
        assert_eq!(&bytes[..8], &[0, 0x61, 0x73, 0x6d, 0x0d, 0, 1, 0]);
        assert!(
            vybe_platform_wasm::write_artifact(
                &compiled.chunks,
                false,
                compiled.component_binary.as_deref(),
                Default::default(),
            )
            .unwrap_err()
            .contains("missing its source marker")
        );
        assert!(
            vybe_platform_wasm::write_artifact(
                &compiled.chunks,
                compiled.component_source,
                compiled.component_binary.as_deref(),
                vybe_platform_wasm::WasmWriteOptions {
                    include_vybe_metadata: true,
                    include_name_section: false,
                },
            )
            .unwrap_err()
            .contains("does not support Vybe metadata")
        );

        let unsupported = super::parse("(component (type $t (component)))")
            .expect("unsupported binary shape still parses for execution");
        let compiled = vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile_with_imports(&unsupported)
            .expect("unsupported binary shape still compiles for execution");
        assert!(compiled.component_source);
        assert!(compiled.component_binary.is_none());
        assert!(
            vybe_platform_wasm::write_artifact(
                &compiled.chunks,
                compiled.component_source,
                compiled.component_binary.as_deref(),
                Default::default(),
            )
            .unwrap_err()
            .contains("does not support all source definitions")
        );

        let core = super::parse("(module (func))").expect("core WAT parses");
        let compiled = vybe_compiler::primitives::Compiler::with_profile(
            vybe_compiler::profile::parse_profile(super::profile_source()).expect("WAST profile"),
        )
        .compile_with_imports(&core)
        .expect("core WAT compiles");
        let bytes = vybe_platform_wasm::write_artifact(
            &compiled.chunks,
            compiled.component_source,
            compiled.component_binary.as_deref(),
            Default::default(),
        )
        .expect("core WAT emits a core module");
        assert_eq!(&bytes[..8], &[0, 0x61, 0x73, 0x6d, 1, 0, 0, 0]);
    }

    #[test]
    fn unsupported_component_definition_is_not_silently_encoded() {
        let module = super::parse("(component (type $t (component)))").expect("component parses");
        assert!(module.canon.component_source);
        assert!(module.canon.binary.is_none());
        assert_eq!(module.canon.types.len(), 1);

        let module = super::parse("(component (core type $mt (module)))")
            .expect("core module type parses");
        assert!(module.canon.component_source);
        assert!(module.canon.binary.is_none());

        let module = super::parse("(module (func))").expect("core module parses");
        assert!(!module.canon.component_source);
    }

    #[test]
    fn component_with_memory_exports_normalizes_both_forms() {
        for source in [
            include_str!(
                "../../../tests/wast/wat_component/test_core_instance_with_imports_memory.wat"
            ),
            include_str!(
                "../../../tests/wast/wat_component/test_core_instance_with_whole_memory.wat"
            ),
        ] {
            let module = super::parse(source).expect("core instance memory import should resolve");
            assert!(!module.body.is_empty());
        }
    }

    #[test]
    fn canonical_dynamic_value_options_follow_lift_and_lower_direction() {
        let valid = include_str!(
            "../../../tests/wast/wat_component/test_canon_lift_requires_declared_realloc.wat"
        );
        super::parse(valid).expect("declared allocator is valid");
        let missing = valid.replace("(realloc (core func $realloc))", "");
        let error = super::parse(&missing)
            .expect_err("lifting a string parameter needs a declared allocator");
        assert!(
            error.contains("canonical option `realloc` is required"),
            "{error}"
        );

        let nested = missing.replace(
            "(type $ft (func (param \"text\" string) (result u32)))",
            "(type $text (record (field \"value\" string)))\n  \
             (type $ft (func (param \"text\" $text) (result u32)))",
        );
        let error = super::parse(&nested)
            .expect_err("a named nested string field needs the same allocator");
        assert!(
            error.contains("canonical option `realloc` is required"),
            "{error}"
        );

        let duplicate = valid.replace(
            "(realloc (core func $realloc))",
            "(realloc (core func $realloc)) (realloc (core func $realloc))",
        );
        let error = super::parse(&duplicate)
            .expect_err("a duplicate canonical option must not replace the first");
        assert!(error.contains("canonical option `realloc` is specified more than once"), "{error}");

        let lower = include_str!(
            "../../../tests/wast/wat_component/test_canon_lower_result_calls_declared_realloc.wat"
        );
        let lower_missing = lower.replace(
            "(canon lower (func $lifted) (memory (core memory $mi \"mem\"))\n               (realloc (core func $r))",
            "(canon lower (func $lifted) (memory (core memory $mi \"mem\"))",
        );
        assert_ne!(lower_missing, lower);
        let error = super::parse(&lower_missing)
            .expect_err("lowering a function with a string result needs realloc for lifting");
        assert!(
            error.contains("canonical option `realloc` is required"),
            "{error}"
        );
    }

    #[test]
    fn canonical_memory_alias_uses_its_own_index_space() {
        let module = super::parse(include_str!(
            "../../../tests/wast/wat_component/test_alias_core_memory_for_canon.wat"
        ))
        .expect("core memory alias should resolve");
        assert_eq!(module.canon.defs.len(), 2);
        assert_eq!(module.canon.defs[0].opts.memory, Some(0));
        assert_eq!(module.canon.defs[1].opts.memory, Some(0));

        let profile =
            vybe_compiler::profile::parse_profile(super::profile_source()).expect("WAST profile");
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile(&module)
            .expect("aliased core memory should compile");
        let binary = vybe_platform_wasm::write_wasm(&chunks);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary)
            .expect("aliased memory fixture must emit standard Wasm");
        let mut vm = vybe_runtime::VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        vm.run(chunks)
            .expect("aliased memory fixture should execute");

        let error = super::parse(
            "(component (canon waitable-set.poll (memory (core memory 0))
                (core func $poll)))",
        )
        .expect_err("an undeclared core memory index must not resolve to VM memory 0");
        assert!(error.contains("memory index out of bounds"), "{error}");

        let shifted = super::parse(
            "(component
               (core module $first (memory (export \"mem\") 1))
               (core instance $a (instantiate $first))
               (core module $second
                 (memory (export \"mem\") 1)
                 (func (export \"id\") (param i32) (result i32) (local.get 0)))
               (core instance $b (instantiate $second))
               (alias core export $b \"mem\" (core memory $mem))
               (type $ft (func (param \"value\" u32) (result u32)))
               (canon lift (core func $b \"id\") (memory $mem)
                 (func $id (type $ft))))",
        )
        .expect("aliased memory index must resolve independently of script slot");
        assert_eq!(shifted.canon.defs[0].opts.memory, Some(1));
    }

    #[test]
    fn canonical_memory_resolves_core_instance_export() {
        let module = super::parse(include_str!(
            "../../../tests/wast/wat_component/test_plain_record_list_canonical_roundtrip.wat"
        ))
        .expect("named core memory export should resolve");
        assert_eq!(module.canon.defs.len(), 2);
        assert_eq!(module.canon.defs[0].opts.memory, Some(0));
        assert_eq!(module.canon.defs[1].opts.memory, Some(0));

        let profile =
            vybe_compiler::profile::parse_profile(super::profile_source()).expect("WAST profile");
        let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
            .compile(&module)
            .expect("canonical record-list round-trip should compile");
        let binary = vybe_platform_wasm::write_wasm(&chunks);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&binary)
            .expect("round-trip must emit standard Wasm");
        let mut vm = vybe_runtime::VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        vm.run(chunks)
            .expect("record-list round-trip should execute");

        let error = super::parse(
            "(component
               (core module $m (memory (export \"mem\") 1))
               (core instance $i (instantiate $m))
               (canon waitable-set.poll (memory (core memory $i \"missing\"))
                 (core func $poll)))",
        )
        .expect_err("an unknown memory export must not fall back to memory zero");
        assert!(error.contains("exports no memory `missing`"), "{error}");
    }

    #[test]
    fn inline_core_func_lower_uses_canonical_function_space() {
        let module = super::parse(
            "(component
               (import \"f\" (func $f (param \"x\" u32)))
               (core func $lowered (canon lower (func $f)))
               (core instance (export \"f\" (func $lowered))))",
        )
        .expect("inline canon lower should define a core function");
        assert_eq!(module.canon.defs.len(), 1);
        assert_eq!(module.canon.defs[0].builtin, "lower");
        assert_eq!(module.canon.defs[0].binder.as_deref(), Some("lowered"));
        assert!(matches!(
            &module.canon.defs[0].callee,
            Some(vybe_ast::canon::CanonCallee::Component(0))
        ));

        let unnamed = super::parse(
            "(component
               (import \"f\" (func $f (param \"x\" u32)))
               (core func (canon lower (func $f)))
               (core func $second (canon lower (func $f)))
               (core instance (export \"first\" (func 0))
                              (export \"second\" (func $second))))",
        )
        .expect("unnamed inline lower must occupy a core function index");
        assert_eq!(unnamed.canon.defs.len(), 2);
        assert_eq!(unnamed.canon.defs[0].binder, None);
        assert_eq!(unnamed.canon.defs[1].binder.as_deref(), Some("second"));
    }

    #[test]
    fn inline_core_export_can_name_an_instantiated_module_export() {
        super::parse(include_str!(
            "../../../tests/wast/wat_component/test_resource_destructor_rep.wat"
        ))
        .expect("core instance export should resolve by instance and export name");
    }

    #[test]
    fn resource_type_retains_named_core_destructor() {
        let module = super::parse(
            "(component
               (core module $M (func (export \"drop\") (param i32)))
               (core instance $m (instantiate $M))
               (type $r (resource (rep i32) (dtor (core func $m \"drop\"))))
               (canon resource.drop $r (core func $drop)))",
        )
        .expect("resource destructor should parse");
        assert!(matches!(
            &module.canon.types[0],
            vybe_ast::canon::TypeDecl::Resource {
                name: Some(name),
                dtor: Some(vybe_ast::canon::CanonCallee::CoreExport { method, .. }),
                ..
            } if name == "r" && method == "drop"
        ));
        let wide = super::parse("(component (type $r (resource (rep i64))))")
            .expect("i64 resource representation should parse");
        assert!(matches!(
            wide.canon.types[0],
            vybe_ast::canon::TypeDecl::Resource {
                rep: vybe_ast::canon::ResourceRep::I64,
                ..
            }
        ));
    }

    #[test]
    fn parses_typed_stream_and_future_forward_declarations() {
        let module = super::parse(
            "(component
               (type $stream (stream u32))
               (type $future (future u32))
               (canon stream.forward $stream (core func $sf))
               (canon future.forward $future (core func $ff)))",
        )
        .expect("forward declarations should parse");
        assert_eq!(module.canon.defs.len(), 2);
        assert_eq!(module.canon.defs[0].builtin, "stream.forward");
        assert_eq!(module.canon.defs[0].ty, Some(0));
        assert_eq!(module.canon.defs[1].builtin, "future.forward");
        assert_eq!(module.canon.defs[1].ty, Some(1));
    }

    #[test]
    fn component_definition_spelling_matches_component() {
        let plain = super::parse("(component (type $t u32))").expect("plain component");
        let explicit = super::parse("(component definition (type $t u32))")
            .expect("explicit component definition");
        assert_eq!(explicit.canon.types, plain.canon.types);
    }

    #[test]
    fn defined_type_labels_are_validated_after_parsing() {
        use vybe_ast::StmtKind;

        for (body, reason) in [
            (
                "(record (field \"\" s32))",
                "record field name cannot be empty",
            ),
            (
                "(variant (case \"\" s32))",
                "variant case name cannot be empty",
            ),
            ("(flags \"\")", "flag name cannot be empty"),
            ("(enum \"\")", "enum tag name cannot be empty"),
            (
                "(func (param \"\" string))",
                "function parameter name cannot be empty",
            ),
            ("(record)", "record type must have at least one field"),
            ("(variant)", "variant type must have at least one case"),
            ("(flags)", "flags must have at least one entry"),
            ("(enum)", "enum type must have at least one variant"),
            ("(tuple)", "tuple type must have at least one type"),
            (
                "(record (field \"a-B-c-D\" string) (field \"A-b-C-d\" u8))",
                "record field name `A-b-C-d` conflicts with previous field name `a-B-c-D`",
            ),
            (
                "(variant (case \"x\" s64) (case \"X\" s64))",
                "variant case name `X` conflicts with previous case name `x`",
            ),
            (
                "(flags \"x\" \"X\")",
                "flag name `X` conflicts with previous flag name `x`",
            ),
            (
                "(enum \"x\" \"X\")",
                "enum tag name `X` conflicts with previous tag name `x`",
            ),
            (
                "(func (param \"foo\" string) (param \"FOO\" u32))",
                "function parameter name `FOO` conflicts with previous parameter name `foo`",
            ),
            (
                "(enum \"NevEr\")",
                "enum tag name `NevEr` is not in kebab case",
            ),
            (
                "(record (field \"GoNnA\" string))",
                "record field name `GoNnA` is not in kebab case",
            ),
            (
                "(variant (case \"GIVe\" string))",
                "variant case name `GIVe` is not in kebab case",
            ),
            (
                "(func (param \"yOu\" string))",
                "function parameter name `yOu` is not in kebab case",
            ),
        ] {
            let source = format!("(assert_invalid (component (type {body})) \"{reason}\")");
            let module = super::parse(&source).expect("invalid type should parse");
            assert!(
                !module
                    .body
                    .iter()
                    .any(|stmt| matches!(stmt.kind, StmtKind::Throw { .. })),
                "{source}: {:?}",
                module.body,
            );
        }

        let wrong_reason = super::parse(
            "(assert_invalid (component (type (record))) \"enum type must have at least one variant\")",
        )
        .expect("assertion should parse");
        assert!(
            wrong_reason
                .body
                .iter()
                .any(|stmt| matches!(stmt.kind, StmtKind::Throw { .. }))
        );

        let flags = |count| {
            let labels = (1..=count)
                .map(|n| format!("\"f{n}\""))
                .collect::<Vec<_>>()
                .join(" ");
            format!("(component (type (flags {labels})))")
        };
        super::parse(&flags(32)).expect("32 flags are valid");
        let too_many = format!(
            "(assert_invalid {} \"cannot have more than 32 flags\")",
            flags(33)
        );
        let module = super::parse(&too_many).expect("33 flags should parse for validation");
        assert!(
            !module
                .body
                .iter()
                .any(|stmt| matches!(stmt.kind, StmtKind::Throw { .. }))
        );
    }

    #[test]
    fn core_module_import_type_ascription_preserves_sort_and_index() {
        use vybe_ast::StmtKind;

        let valid = super::parse(
            "(component
                (core type $mt (module))
                (import \"m\" (core module $imported (type $mt)))
                (core module $local (func (export \"f\")))
                (core instance $instance (instantiate $local)))",
        )
        .expect("a typed import should not displace the following local module");
        assert!(
            !valid
                .body
                .iter()
                .any(|stmt| matches!(stmt.kind, StmtKind::Throw { .. }))
        );

        for (type_decl, expected) in [
            (
                "(core type $t (func))",
                "core type index 0 is not a module type",
            ),
            ("", "core type index out of bounds"),
        ] {
            let source = format!(
                "(assert_invalid (component {type_decl} (import \"m\" (core module (type 0)))) \"{expected}\")"
            );
            let module = super::parse(&source).expect("invalid import should parse");
            assert!(
                !module
                    .body
                    .iter()
                    .any(|stmt| matches!(stmt.kind, StmtKind::Throw { .. })),
                "{source}: {:?}",
                module.body,
            );
        }

        let missing = super::parse(
            "(component
                (core type $mt (module))
                (import \"m\" (core module $m (type $mt)))
                (core instance (instantiate $m)))",
        )
        .expect_err("an unresolved imported module cannot be instantiated");
        assert!(missing.contains("has no linked definition"), "{missing}");
    }

    #[test]
    fn value_type_references_require_prior_defined_types() {
        use vybe_ast::StmtKind;

        super::parse(
            "(component
                (type $base (record (field \"x\" u8)))
                (type (option $base))
                (type (func (param \"x\" (list $base)) (result $base))))",
        )
        .expect("prior value types are valid in nested positions");

        for (prefix, body, reason) in [
            (
                "(type $t (func))",
                "(func (param \"x\" $t))",
                "type index 0 is not a defined type",
            ),
            (
                "(type $t (instance))",
                "(func (result $t))",
                "type index 0 is not a defined type",
            ),
            (
                "(type $t (component))",
                "(option $t)",
                "type index 0 is not a defined type",
            ),
            ("", "(option 0)", "type index out of bounds"),
            ("", "(record (field \"x\" 0))", "type index out of bounds"),
            ("", "(result 0 (error 1))", "type index out of bounds"),
        ] {
            let source =
                format!("(assert_invalid (component {prefix} (type {body})) \"{reason}\")");
            let module = super::parse(&source).expect("invalid type should parse");
            assert!(
                !module
                    .body
                    .iter()
                    .any(|stmt| matches!(stmt.kind, StmtKind::Throw { .. })),
                "{source}: {:?}",
                module.body,
            );
        }
    }

    #[test]
    fn function_type_ascriptions_check_bounds_and_sort() {
        use vybe_ast::StmtKind;

        super::parse(
            "(component
                (type $ft (func))
                (import \"f\" (func $f (type $ft)))
                (core module $m (func (export \"g\")))
                (core instance $i (instantiate $m))
                (func (type $ft) (canon lift (core func $i \"g\"))))",
        )
        .expect("valid function type indices should be accepted");

        for (prefix, import, reason) in [
            (
                "",
                "(import \"f\" (func (type 100)))",
                "type index out of bounds",
            ),
            (
                "(type $t (instance))",
                "(import \"f\" (func (type $t)))",
                "type index 0 is not a function type",
            ),
            (
                "(type $t (record (field \"x\" u8)))",
                "(import \"f\" (func (type $t)))",
                "type index 0 is not a function type",
            ),
        ] {
            let source = format!("(assert_invalid (component {prefix} {import}) \"{reason}\")");
            let module = super::parse(&source).expect("invalid import should parse");
            assert!(
                !module
                    .body
                    .iter()
                    .any(|stmt| matches!(stmt.kind, StmtKind::Throw { .. })),
                "{source}: {:?}",
                module.body,
            );
        }

        let lift = super::parse(
            "(assert_invalid
                (component
                  (core module $m (func (export \"f\")))
                  (core instance $i (instantiate $m))
                  (func (type 100) (canon lift (core func $i \"f\"))))
                \"type index out of bounds\")",
        )
        .expect("invalid lift type should parse");
        assert!(
            !lift
                .body
                .iter()
                .any(|stmt| matches!(stmt.kind, StmtKind::Throw { .. }))
        );
    }

    #[test]
    fn component_assert_invalid_checks_the_actual_reason() {
        use vybe_ast::StmtKind;

        let matching = super::parse(
            "(assert_invalid
                (component (import \"x\" (func)) (import \"x\" (func)))
                \"conflicts with previous name\")",
        )
        .expect("component assertion should parse");
        assert!(
            !matching
                .body
                .iter()
                .any(|stmt| matches!(stmt.kind, StmtKind::Throw { .. })),
            "{:?}",
            matching.body,
        );

        for source in [
            "(assert_invalid (component (type u32)) \"conflicts with previous name\")",
            "(assert_invalid
                (component (import \"x\" (func)) (import \"x\" (func)))
                \"exceeds maximum byte size\")",
        ] {
            let module = super::parse(source).expect("component assertion should parse");
            assert!(
                module
                    .body
                    .iter()
                    .any(|stmt| matches!(stmt.kind, StmtKind::Throw { .. }))
            );
        }
    }

    #[test]
    fn nested_core_module_is_validated_without_instantiation() {
        use vybe_ast::StmtKind;

        let invalid = super::parse(
            "(assert_invalid (component (core module (func i32.add))) \"type mismatch\")",
        )
        .expect("nested core module should parse");
        assert!(
            !invalid
                .body
                .iter()
                .any(|stmt| matches!(stmt.kind, StmtKind::Throw { .. })),
            "{:?}",
            invalid.body,
        );

        let valid = super::parse(
            "(component definition (core module $m (func (export \"f\") (result i32) i32.const 1)))",
        )
        .expect("valid nested core module should parse");
        assert!(
            !valid
                .body
                .iter()
                .any(|stmt| matches!(stmt.kind, StmtKind::Throw { .. })),
            "{:?}",
            valid.body,
        );
    }

    #[test]
    fn core_module_type_checks_its_own_declarations() {
        use vybe_ast::StmtKind;

        for (body, expected) in [
            ("(export \"a\" (func (type 0)))", "type index out of bounds"),
            (
                "(export \"a\" (func (type $missing)))",
                "type index out of bounds",
            ),
            (
                "(export \"a\" (func (type 0))) (type (func))",
                "type index out of bounds",
            ),
            (
                "(export \"a\" (func)) (export \"a\" (func))",
                "export name `a` already defined",
            ),
            (
                "(import \"\" \"\" (memory 70000))",
                "memory size must be at most",
            ),
            (
                "(import \"\" \"a\" (func)) (import \"\" \"a\" (func))",
                "duplicate import name `:a`",
            ),
        ] {
            let source =
                format!("(assert_invalid (component (core type (module {body}))) \"{expected}\")");
            let module = super::parse(&source).expect("core module type should parse");
            assert!(
                !module
                    .body
                    .iter()
                    .any(|stmt| matches!(stmt.kind, StmtKind::Throw { .. })),
                "{source}: {:?}",
                module.body,
            );
        }
        let valid = super::parse(
            "(component (core type (module
                (type $f (func))
                (import \"a\" \"b\" (func (type $f)))
                (export \"e\" (func (type $f))))))",
        )
        .expect("valid named core module type should parse");
        assert!(
            !valid
                .body
                .iter()
                .any(|stmt| matches!(stmt.kind, StmtKind::Throw { .. })),
            "{:?}",
            valid.body,
        );
        let implicit = super::parse(
            "(component (core type (module
                (import \"a\" \"b\" (func))
                (export \"c\" (func (type 0))))))",
        )
        .expect("implicit core function type should parse");
        assert!(
            !implicit
                .body
                .iter()
                .any(|stmt| matches!(stmt.kind, StmtKind::Throw { .. })),
            "{:?}",
            implicit.body,
        );
    }

    #[test]
    fn component_extern_names_validate_and_fold_hyphens() {
        use vybe_ast::StmtKind;

        for source in [
            "(assert_invalid (component (import \"wasi/http\" (func))) \"`wasi/http` is not in kebab case\")",
            "(assert_invalid (component (import \"wasi:\" (func))) \"`` is not in kebab case\")",
            "(assert_invalid (component (import \"foo-bar\" (func)) (import \"foobar\" (func))) \"import name `foobar` conflicts with previous name `foo-bar`\")",
            "(assert_invalid (component (type (component (import \"GonnA\" (func))))) \"`GonnA` is not in kebab case\")",
            "(assert_invalid (component (type (instance (export \"lET\" (func))))) \"`lET` is not in kebab case\")",
            "(assert_invalid (component (type (component (export \"NevEr\" (func))))) \"`NevEr` is not in kebab case\")",
            "(assert_invalid (component (type (component (import \"foo-bar\" (func)) (import \"foobar\" (func))))) \"import name `foobar` conflicts with previous name `foo-bar`\")",
            "(assert_invalid (component (import \"f\" (func $f)) (instance (export \"1\" (func $f)))) \"`1` is not in kebab case\")",
            "(assert_invalid (component (import \"f\" (func $f)) (instance (export \"foo-bar\" (func $f)) (export \"foobar\" (func $f)))) \"export name `foobar` conflicts with previous name `foo-bar`\")",
        ] {
            let module = super::parse(source).expect("component assertion should parse");
            assert!(
                !module
                    .body
                    .iter()
                    .any(|stmt| matches!(stmt.kind, StmtKind::Throw { .. })),
                "{source}: {:?}",
                module.body,
            );
        }
    }

    #[test]
    fn component_interface_names_validate_semver_and_current_nesting() {
        use vybe_ast::StmtKind;

        for name in [
            "wasi:http/types",
            "wasi:http/types@1.0.0",
            "a:b/c@0.0.0-abcd.1+efg.4",
        ] {
            super::parse(&format!("(component (import \"{name}\" (func)))"))
                .expect("valid interface name");
        }
        for (name, expected) in [
            ("a:b/c@", "empty string"),
            ("a:b/c@.", "unexpected character '.'"),
            ("a:b/c@1.", "unexpected end of input"),
            ("a:b/c@2.0.0+", "empty identifier segment"),
            ("foo:bar:baz/qux", "expected `/` after package name"),
            ("foo:bar/baz/qux", "trailing characters found: `/qux`"),
        ] {
            let source =
                format!("(assert_invalid (component (import \"{name}\" (func))) \"{expected}\")");
            let module = super::parse(&source).expect("invalid interface name should parse");
            assert!(
                !module
                    .body
                    .iter()
                    .any(|stmt| matches!(stmt.kind, StmtKind::Throw { .. })),
                "{source}: {:?}",
                module.body,
            );
        }
    }
}
