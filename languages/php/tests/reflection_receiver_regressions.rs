//! Regression coverage for PHP reflection's explicit receiver ABI.
use vybe_runtime::{Value, VM};

fn run(source: &str) -> Value {
    static INIT: std::sync::Once = std::sync::Once::new();
    INIT.call_once(vybe_language_php::register);
    let ast = vybe_language_php::parse(source).unwrap();
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
    let chunks = vybe_compiler::Compiler::with_profile(profile)
        .compile(&ast)
        .unwrap();
    let helper = chunks
        .iter()
        .find(|chunk| chunk.name == "__refl_implements")
        .expect("reflection interface helper emitted");
    assert!(
        helper.takes_receiver,
        "local 0 must receive the reflection object"
    );
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vm.run(chunks).unwrap()
}

#[test]
fn dynamic_reflection_interface_method_receives_its_object() {
    assert_eq!(
        run(r#"<?php
            interface AdapterContract {}
            class AdapterImplementation implements AdapterContract {}
            $name = AdapterImplementation::class;
            $reflection = new ReflectionClass($name);
            return $reflection->implementsInterface(AdapterContract::class);
        "#),
        Value::Bool(true)
    );
}

#[test]
fn dynamic_reflection_does_not_accept_unrelated_interface() {
    assert_eq!(
        run(r#"<?php
            interface AdapterContract {}
            interface OtherContract {}
            class AdapterImplementation implements AdapterContract {}
            $name = AdapterImplementation::class;
            $reflection = new ReflectionClass($name);
            return $reflection->implementsInterface(OtherContract::class);
        "#),
        Value::Bool(false)
    );
}

#[test]
fn dynamic_reflection_name_getter_receives_its_object() {
    assert_eq!(
        run(r#"<?php
            interface AdapterContract {}
            class AdapterImplementation implements AdapterContract {}
            $name = AdapterImplementation::class;
            $reflection = new ReflectionClass($name);
            return $reflection->getName();
        "#),
        Value::String("AdapterImplementation".into())
    );
}
