use std::sync::Once;

fn register_powershell() {
    static R: Once = Once::new();
    R.call_once(vybe_language_powershell::register);
}

#[test]
fn powershell_profile_extensions_include_ps1() {
    register_powershell();
    let lang = vybe_compiler::languages::find_by_name("powershell")
        .expect("powershell language not registered");
    let src = (lang.profile_source)();
    assert!(src.contains("ps\""));
    assert!(src.contains("ps1"));
    assert!(src.contains("psm1"));
    assert!(src.contains("psd1"));
}

#[test]
fn powershell_parse_smoke() {
    register_powershell();
    let ast = vybe_language_powershell::parse("Write-Output hello\n");
    assert!(
        ast.is_ok(),
        "parser should accept basic command text: {:?}",
        ast.err()
    );
}

#[test]
fn powershell_extension_lookup_supports_ps() {
    register_powershell();
    assert!(vybe_compiler::languages::find_by_extension("ps").is_some());
}
