//! Parsed PHP regressions for the shared dynamic formatter and fractional constructor.
use vybe_runtime::{Chunk, VM, Value};
fn compile(source: &str) -> Vec<Chunk> {
    static INIT: std::sync::Once = std::sync::Once::new();
    INIT.call_once(vybe_language_php::register);
    let ast = vybe_language_php::parse(&format!("<?php {source}")).unwrap();
    let profile = vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
    vybe_compiler::Compiler::with_profile(profile).compile(&ast).unwrap()
}
fn run(source: &str, expected: &str) {
    let chunks = compile(source);
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    assert_eq!(vm.run(chunks).unwrap(), Value::String(expected.into()));
    assert!(vm.last_exception.is_none());
}
#[test]
fn dynamic_formatter_preserves_calendar_tokens_and_small_emission() {
    for (format, expected) in [("Y-m-d H:i:s", "2004-02-29 00:05:06"),("YmdHis", "20040229000506"),("c", "2004-02-29T00:05:06+00:00"),("u a A", "123456 am AM"),("\\Y Y ! ?", "Y 2004 ! ?")] {
        let source = format!("$format='{format}';$d=new DateTimeImmutable('2004-02-29 00:05:06.123456',new DateTimeZone('UTC'));return $d->format($format);");
        let chunks = compile(&source);
        let helper = chunks.iter().find(|c| c.name == "__php_datetime_format_runtime").expect("actual dynamic formatter");
        assert!(helper.code.len()<19_000, "formatter should retain its specialized emission and first-match dispatch, got {} bytes", helper.code.len());
        run(&source, expected);
    }
}
#[test]
fn fractional_constructor_retains_dynamic_and_literal_microseconds() {
    for (digits, expected) in [("1","100000"),("12","120000"),("123","123000"),("1234","123400"),("12345","123450"),("123456","123456")] {
        run(&format!("$text='2004-02-29 00:05:06.{digits}';$d=new DateTimeImmutable($text,new DateTimeZone('UTC'));$format='u';return $d->format($format);"), expected);
    }
    run("$d=new DateTimeImmutable('2004-02-29 00:05:06.123456');$format='u';return $d->format($format);", "123456");
}
