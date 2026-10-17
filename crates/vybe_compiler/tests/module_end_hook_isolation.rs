use vybe_compiler::primitives::Compiler;
use vybe_runtime::{Chunk, Value};

fn mentions(chunks: &[Chunk], name: &str) -> bool {
    chunks.iter().any(|chunk| {
        chunk.globals.iter().any(|global| global == name)
            || chunk
                .constants
                .iter()
                .any(|value| matches!(value, Value::String(s) if s.as_ref() == name))
    })
}

#[test]
fn c_modules_do_not_acquire_php_shutdown_or_buffer_epilogues() {
    let module = vybe_language_c::parse("int main(void) { return 0; }").unwrap();
    assert!(module.directives.shutdown_callbacks.is_none());
    assert_ne!(module.directives.flush_output_buffers, Some(true));
    let profile = vybe_compiler::profile::parse_profile(vybe_language_c::profile_source()).unwrap();
    let chunks = Compiler::with_profile(profile).compile(&module).unwrap();
    assert!(!mentions(&chunks, "__php_shutdown_fns"));
    assert!(!mentions(&chunks, "__vybe_ob_stack"));
    assert!(
        !chunks
            .iter()
            .flat_map(|c| &c.imports)
            .any(|i| (i.module == "ecma:array" && i.name == "slice")
                || (i.module == "ecma:function" && i.name == "apply"))
    );
}

#[test]
fn php_programs_keep_hooks_but_included_units_defer_them() {
    let module = vybe_language_php::parse("<?php echo 'ok';").unwrap();
    assert!(matches!(
        module.directives.shutdown_callbacks,
        Some(vybe_ast::ShutdownPolicy::Callbacks { .. })
    ));
    assert_eq!(module.directives.flush_output_buffers, Some(true));
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
    let program = Compiler::with_profile(profile.clone())
        .compile(&module)
        .unwrap();
    assert!(mentions(&program, "__php_shutdown_fns"));
    let included = Compiler::with_profile(profile)
        .with_run_module_end_hooks(false)
        .compile(&module)
        .unwrap();
    assert!(!mentions(&included, "__php_shutdown_fns"));
}

#[test]
fn unspecified_lifecycle_policy_inherits_explicit_module_policy() {
    let mut policy = vybe_ast::Directives {
        shutdown_callbacks: Some(vybe_ast::ShutdownPolicy::Callbacks {
            registry: "custom.callbacks".into(),
            order: vybe_ast::ShutdownOrder::Registration,
        }),
        flush_output_buffers: Some(true),
        ..Default::default()
    };
    policy.overlay(&Default::default());
    assert!(matches!(
        policy.shutdown_callbacks,
        Some(vybe_ast::ShutdownPolicy::Callbacks { .. })
    ));
    assert_eq!(policy.flush_output_buffers, Some(true));
    policy.overlay(&vybe_ast::Directives {
        shutdown_callbacks: Some(vybe_ast::ShutdownPolicy::Disabled),
        flush_output_buffers: Some(false),
        ..Default::default()
    });
    assert_eq!(
        policy.shutdown_callbacks,
        Some(vybe_ast::ShutdownPolicy::Disabled)
    );
    assert_eq!(policy.flush_output_buffers, Some(false));
}

#[test]
fn shutdown_registry_and_order_come_from_ast_not_language_name() {
    use vybe_ast::{ShutdownOrder, ShutdownPolicy};
    for (order, expected) in [
        (ShutdownOrder::Registration, 12.0),
        (ShutdownOrder::ReverseRegistration, 21.0),
    ] {
        let mut module = vybe_language_php::parse(
            "<?php\n\
            $trace = 0;\n\
            function first() { global $trace; $trace = $trace * 10 + 1; }\n\
            function second() { global $trace; $trace = $trace * 10 + 2; }\n\
            register_shutdown_function('first');\n\
            register_shutdown_function('second');",
        )
        .unwrap();
        module.directives.shutdown_callbacks = Some(ShutdownPolicy::Callbacks {
            registry: "custom.callbacks".into(),
            order,
        });
        module.directives.flush_output_buffers = Some(false);
        let profile =
            vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
        let mut renamed = profile.clone();
        renamed.name = "another-language".into();
        let chunks = Compiler::with_profile(renamed).compile(&module).unwrap();
        assert!(mentions(&chunks, "custom.callbacks"));
        assert!(!mentions(&chunks, "__php_shutdown_fns"));
        // Runtime variable bindings remain those of the source frontend;
        // termination registry/order are independently declared in the AST.
        let chunks = Compiler::with_profile(profile).compile(&module).unwrap();
        let mut vm = vybe_runtime::VM::new();
        vybe_compiler::primitives::platforms::register_platforms(
            &mut vm,
            &vybe_runtime::capabilities::Capabilities::all(),
        );
        let script = vybe_runtime::debug::disassemble(&chunks[0]);
        if let Err(error) = vm.run(chunks) {
            panic!("{order:?}: {error:?}\n{script}");
        }
        let Some(Value::Object(global_this)) = vm.global("globalThis") else {
            panic!("request global object is missing");
        };
        let store = global_this.lock().unwrap().properties.get("__vybe_php_globals")
            .cloned().expect("request variable store is missing");
        let Value::Object(store) = store else { panic!("request store is not an object"); };
        let trace = store.lock().unwrap().properties.get("trace").cloned();
        assert!(matches!(trace, Some(Value::F64(n)) if n == expected), "{order:?}: {trace:?}");
    }
}
