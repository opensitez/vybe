use vybe_runtime::{VM, Value};

fn repository_root() -> &'static std::path::Path {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .ancestors()
        .find(|root| root.join("examples/vb/countries/mysqlcountries.vbproj").is_file())
        .expect("test manifest is within the Vybe repository");
    static CWD: std::sync::Once = std::sync::Once::new();
    CWD.call_once(|| std::env::set_current_dir(root).unwrap());
    root
}

#[test]
fn winforms_binding_source_and_grid_behavior_regressions() {
    vybe_language_vb::register();
    vybe_platform_dotnet::register();
    let fixtures = repository_root().join("tests/vb/vb_winforms_data_binding");
    for file in std::fs::read_dir(fixtures).unwrap() {
        let path = file.unwrap().path();
        if path.extension().and_then(|x| x.to_str()) != Some("vb") { continue; }
        let bundle = vybe_compiler::projects::load(&path).unwrap();
        let mut vm = VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        let module = bundle.prepared_module().unwrap();
        let compiled = bundle.compile_prepared_module(&module, &vm.modules).unwrap();
        let result = vm.run(compiled.chunks);
        assert!(result.is_ok(), "{}: {result:?}", path.display());
        assert!(vm.last_exception.is_none(), "{}: guest exception", path.display());
        eprintln!("passed {}", path.display());
    }
}

#[test]
fn actual_countries_project_initializes_all_three_control_bindings() {
    vybe_language_vb::register();
    vybe_platform_dotnet::register();
    let project = repository_root().join("examples/vb/countries/mysqlcountries.vbproj");
    let bundle = vybe_compiler::projects::load(&project).unwrap();
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let module = bundle.prepared_module().unwrap();
    let compiled = bundle.compile_prepared_module(&module, &vm.modules).unwrap();
    vm.run(compiled.chunks).unwrap();
    assert!(vm.last_exception.is_none());
    for (global, control) in [
        ("__dotnet_bindingnavigator_sources", "bnav1"),
        ("__dotnet_grid_sources", "dgv1"),
    ] {
        let Value::Object(registry) = vm.global(global).unwrap() else {
            panic!("{global} must initialize an Object");
        };
        assert!(registry.lock().unwrap().properties.contains_key(control));
    }
    let Value::Object(registry) = vm.global("__dotnet_control_bindings").unwrap() else {
        panic!("control bindings registry must be initialized");
    };
    let registry = registry.lock().unwrap();
    for name in ["txt1", "txt2", "txt3"] {
        let Value::Object(bindings) = registry.properties.get(name).unwrap() else {
            panic!("expected binding list for {name}");
        };
        let bindings = bindings.lock().unwrap();
        let vybe_runtime::value::ObjectKind::Array(values) = &bindings.kind else {
            panic!("expected binding array for {name}");
        };
        assert_eq!(values.len(), 1, "binding added exactly once for {name}");
    }
}

#[test]
fn actual_allcontrols_project_initializes_tab_registry() {
    vybe_language_vb::register();
    vybe_platform_dotnet::register();
    let project = repository_root().join("examples/vb/allcontrols/allcontrols.vbproj");
    let bundle = vybe_compiler::projects::load(&project).unwrap();
    for initial in [None, Some(Value::Null), Some(Value::Undefined)] {
        let mut vm = VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        if let Some(value) = initial {
            vm.set_global("__dotnet_tab_states", value);
        }
        let module = bundle.prepared_module().unwrap();
        let compiled = bundle.compile_prepared_module(&module, &vm.modules).unwrap();
        vm.run(compiled.chunks).unwrap();
        assert!(vm.last_exception.is_none());
        let Value::Object(registry) = vm.global("__dotnet_tab_states").unwrap() else {
            panic!("tab registry must initialize an Object");
        };
        let registry = registry.lock().unwrap();
        let Value::Object(state) = registry.properties.get("tab1").unwrap() else {
            panic!("tab control must retain its state");
        };
        let state = state.lock().unwrap();
        assert!(state.properties.contains_key("page"));
        assert!(state.properties.contains_key("button"));
    }
}
