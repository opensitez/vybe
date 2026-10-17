use vybe_compiler::{dynamic::RuntimeCompilerService, frontend_preparation};
use vybe_runtime::{Chunk, VM};

fn run(chunks: Vec<Chunk>) -> Result<String, (String, String, String)> {
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vm.run(chunks).map(|v| format!("{v:?}")).map_err(|e| {
        (
            e.message.clone(),
            e.to_string(),
            format!("{:?}", e.call_stack),
        )
    })
}

#[test]
fn registered_literal_serialization_preserves_values_and_unknown_fallback() {
    vybe_language_php::register();
    let language = vybe_compiler::languages::find_by_name("php").unwrap();
    let selected = frontend_preparation::snapshot(language.parse).capabilities;
    assert_eq!(selected.argument_emitters.len(), 1);
    assert_eq!(selected.argument_emitters[0].operation, "php.serialize");
    assert!(
        frontend_preparation::snapshot(vybe_language_php::parse_with_peg)
            .capabilities
            .argument_emitters
            .is_empty()
    );
    for (source, literal) in [
        ("return serialize('é😀');", true),
        ("function f(){return serialize('nested');}return f();", true),
        (
            "class C{public function f(){return serialize('method');}}return (new C)->f();",
            true,
        ),
        ("$s='dynamic';return serialize($s);", false),
        ("return serialize([1,2]);", false),
        (
            "function serialize($s){return 'shadow';}return serialize('input');",
            false,
        ),
    ] {
        let compile = |capabilities| {
            frontend_preparation::register(language.parse, capabilities);
            let mut vm = VM::new();
            // Platform/bootstrap registration must retain explicit selection.
            vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
            let mut service = RuntimeCompilerService::new(&mut vm);
            service
                .compile_source(
                    &format!("<?php {source}"),
                    language,
                    "/private/tmp/literal-serialization-regression.php",
                )
                .unwrap()
                .chunks
        };
        let mut disabled = selected;
        disabled.argument_emitters = &[];
        let control = compile(disabled);
        let actual = compile(selected);
        if literal {
            assert!(
                !actual
                    .iter()
                    .any(|c| c.name.starts_with("__php_serialize_value"))
            );
            assert!(
                actual.iter().map(|c| c.code.len()).sum::<usize>()
                    < control.iter().map(|c| c.code.len()).sum::<usize>()
            );
        } else {
            assert_eq!(
                control.iter().map(|c| &c.code).collect::<Vec<_>>(),
                actual.iter().map(|c| &c.code).collect::<Vec<_>>()
            );
            assert_eq!(
                control.iter().map(|c| &c.lines).collect::<Vec<_>>(),
                actual.iter().map(|c| &c.lines).collect::<Vec<_>>()
            );
        }
        assert_eq!(run(control), run(actual));
    }
    frontend_preparation::register(language.parse, selected);
}
