#[test]
fn dump_wasm_imports() {
    let src = r#"
        function main() {
            let a = [10, 20, 30];
            return a.length;
        }
    "#;
    let module = vybe_compiler::languages::js::parse(src).expect("parse");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_compiler::languages::js::profile_source())
            .expect("profile");
    let chunks = vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("compile");
    let wasm = vybe_bytecode::wasm::write_wasm(&chunks);
    std::fs::write("/tmp/test_js.wasm", &wasm).unwrap();
}
