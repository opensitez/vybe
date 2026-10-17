//! PHP syntactic URL splitting must handle unmatched optional regex groups.
use vybe_runtime::{VM, Value};

fn run(source: &str) -> Value {
    static INIT: std::sync::Once = std::sync::Once::new();
    INIT.call_once(vybe_language_php::register);
    let ast = vybe_language_php::parse(source).unwrap();
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
    let chunks = vybe_compiler::Compiler::with_profile(profile)
        .compile(&ast)
        .unwrap();
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vm.run(chunks).unwrap()
}

#[test]
fn absent_query_and_fragment_do_not_reach_string_concat() {
    assert_eq!(
        run(r#"<?php return parse_url('https://example.test/path', PHP_URL_HOST);"#),
        Value::String("example.test".into())
    );
}

#[test]
fn syntactic_split_preserves_case_port_and_dot_segments() {
    assert_eq!(
        run(r#"<?php
            $url = parse_url('HTTP://user:pass@Example.COM:80/a/../b');
            return $url['scheme'].'|'.$url['user'].'|'.$url['pass'].'|'
                .$url['host'].'|'.$url['port'].'|'.$url['path'];
        "#),
        Value::String("HTTP|user|pass|Example.COM|80|/a/../b".into())
    );
}

#[test]
fn present_query_and_fragment_keep_their_content() {
    assert_eq!(
        run(r#"<?php
            $url = parse_url('https://example.test/path?x=1#part');
            return $url['query'].'|'.$url['fragment'];
        "#),
        Value::String("x=1|part".into())
    );
}
