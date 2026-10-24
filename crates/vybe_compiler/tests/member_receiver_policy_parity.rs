//! Isolate receiver-writeback policy from legacy generic string-call lowering.
//! The object-as-string fixture is a common-compiler diagnostic, NOT a native
//! PHP compatibility oracle. Both modes should preserve its returned value.

use vybe_runtime::VM;

fn run(source: &str, writeback: bool) -> (String, String) {
    vybe_language_php::register();
    let mut module = vybe_language_php::parse(&format!("<?php {source}")).unwrap();
    module.directives.member_receiver_writeback = Some(writeback);
    let profile = vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
    let chunks = vybe_compiler::Compiler::with_profile(profile).compile(&module).unwrap();
    let code = chunks.iter().filter(|chunk| chunk.name == "count_subject")
        .map(vybe_runtime::debug::disassemble).collect::<Vec<_>>().join("\n");
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    let outcome = match vm.run(chunks) {
        Ok(value) => format!("Ok({value:?})"),
        Err(error) => format!("Err({})", error.message),
    };
    (outcome, code)
}

#[test]
fn receiver_writeback_does_not_change_legacy_split_return_transport() {
    let source = "class Splitter{function split($needle){return ['a','b','c','d'];}}function count_subject($h,$n){return substr_count($h,$n);}return count_subject(new Splitter(),'x');";
    let (disabled, disabled_code) = run(source, false);
    let (enabled, enabled_code) = run(source, true);
    if disabled != enabled {
        eprintln!("DISABLED {disabled}\n{disabled_code}\nENABLED {enabled}\n{enabled_code}");
    }
    assert_eq!(disabled, enabled, "receiver policy changed return transport");
}
