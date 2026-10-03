#[path = "../generated_tests/build_support/lua_walker.rs"]
mod lua_walker;
fn main() {
    let rustc = std::env::var_os("RUSTC").unwrap_or_else(|| "rustc".into());
    let version = std::process::Command::new(rustc)
        .arg("--version")
        .output()
        .expect("query rustc version");
    assert!(version.status.success());
    println!(
        "cargo:rustc-env=RUSTC_VERSION={}",
        String::from_utf8(version.stdout).unwrap().trim()
    );
    let out = std::path::PathBuf::from(std::env::var_os("OUT_DIR").unwrap());
    lua_walker::generate(&out);
}
