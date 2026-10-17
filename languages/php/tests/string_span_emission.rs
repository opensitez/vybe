//! Regression coverage for PHP string span emitters.
//! These execute emitted code in the VM rather than testing Rust string helpers.

use std::sync::Arc;
use vybe_compiler::Compiler;
use vybe_runtime::{VM, Value};

fn run_cases(cases: &[(&str, &str)]) {
    let mut setup = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut setup);
    vybe_language_php::register();
    let language = vybe_compiler::languages::find_by_name("php").unwrap();
    let profile = Arc::new(
        vybe_runtime::profile::parse_profile((language.profile_source)()).unwrap(),
    );

    for (expression, expected) in cases {
        let source = format!("<?php return {expression};");
        let module = vybe_language_php::parse(&source).unwrap();
        let chunks = Compiler::with_shared_profile(profile.clone())
            .with_run_module_end_hooks(false)
            .compile(&module)
            .unwrap();
        let mut vm = VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        let value = vm.run(chunks).unwrap();
        let Value::String(actual) = value else {
            panic!("{expression} returned a non-string: {value:?}");
        };
        assert_eq!(actual.as_ref(), *expected, "{expression}");
    }
}

#[test]
fn strip_tags_preserves_spans_and_utf16_tag_boundaries() {
    run_cases(&[
        ("strip_tags('')", ""),
        ("strip_tags('plain text')", "plain text"),
        ("strip_tags('0')", "0"),
        ("strip_tags('<b>bold</b> tail')", "bold tail"),
        (
            "strip_tags('before<b>bold</b>after', '<b>')",
            "before<b>bold</b>after",
        ),
        (
            "strip_tags('<B class=\"x\">bold</B><i>italic</i>', '<b>')",
            "<B class=\"x\">bold</B>italic",
        ),
        ("strip_tags('<!-- comment -->kept')", "kept"),
        ("strip_tags('a<br/>b<br>c', '<br>')", "a<br/>b<br>c"),
        (
            "strip_tags('caf\u{e9} <b>na\u{ef}ve</b>')",
            "caf\u{e9} na\u{ef}ve",
        ),
        (
            "strip_tags('\u{1f642} <b>\u{1f680}</b>!')",
            "\u{1f642} \u{1f680}!",
        ),
        (
            "strip_tags('\u{1f642}<b>x</b>\u{1f680}', '<b>')",
            "\u{1f642}<b>x</b>\u{1f680}",
        ),
    ]);
}

#[test]
fn pack_hex_spans_preserve_valid_byte_values() {
    run_cases(&[
        ("bin2hex(pack('H*', ''))", ""),
        ("bin2hex(pack('H*', '00'))", "00"),
        ("bin2hex(pack('H*', '01'))", "01"),
        ("bin2hex(pack('H*', '7f'))", "7f"),
        ("bin2hex(pack('H*', '80'))", "80"),
        ("bin2hex(pack('H*', 'ff'))", "ff"),
        ("bin2hex(pack('H*', '0123456789abcdef'))", "0123456789abcdef"),
        ("bin2hex(pack('H*', 'ABCDEF'))", "abcdef"),
    ]);
}
