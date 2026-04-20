#[test]
fn dump_wasm_imports() {
    let src = r#"
        function main() {
            let a = [10, 20, 30];
            return a.length;
        }
    "#;
    let module = vybex::languages::js::parse(src).expect("parse");
    let profile = vybex::profile::parse_profile(vybex::languages::js::profile_source()).expect("profile");
    let chunks = vybex::compiler::Compiler::with_profile(profile).compile(&module).expect("compile");
    let wasm = vybe_bytecode::wasm::write_wasm(&chunks);
    std::fs::write("/tmp/test_js.wasm", &wasm).unwrap();
}
