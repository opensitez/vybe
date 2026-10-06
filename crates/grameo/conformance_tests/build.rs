#[path = "build_support/javascript_walker.rs"]
mod javascript_walker;
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
    if std::env::var_os("CARGO_FEATURE_JAVASCRIPT_BENCHMARK").is_some() {
        javascript_walker::generate(&out);
        let grammar = "../../../languages/js/src/grammar.pest";
        println!("cargo:rerun-if-changed={grammar}");
        let source = std::fs::read_to_string(grammar).unwrap();
        println!("cargo:rerun-if-changed=fixtures/javascript_before_chain_factor.pest");
        let original =
            std::fs::read_to_string("fixtures/javascript_before_chain_factor.pest").unwrap();
        std::fs::write(
            out.join("javascript_before_factor_parser.rs"),
            grameo_codegen::generate(&original).unwrap(),
        )
        .unwrap();
        println!("cargo:rerun-if-changed=fixtures/javascript_factored.pest");
        let candidate = std::fs::read_to_string("fixtures/javascript_factored.pest").unwrap();
        std::fs::write(
            out.join("javascript_factored_parser.rs"),
            grameo_codegen::generate(&candidate).unwrap(),
        )
        .unwrap();
        std::fs::write(
            out.join("javascript_parser.rs"),
            grameo_codegen::generate(&source).unwrap(),
        )
        .unwrap();
    }
    for (feature, language, start, end, exported) in [
        (
            "CARGO_FEATURE_CORPUS_SCRIPTING",
            "python",
            "fn scan_line_brackets(",
            "\n// ═",
            "preprocess_indentation",
        ),
        (
            "CARGO_FEATURE_CORPUS_BUSINESS",
            "fortran",
            "fn rewrite_hollerith(",
            "\npub fn parse(",
            "rewrite_hollerith",
        ),
    ] {
        if std::env::var_os(feature).is_none() {
            continue;
        }
        let path = format!("../../../languages/{language}/src/walker.rs");
        println!("cargo:rerun-if-changed={path}");
        let source = std::fs::read_to_string(path).unwrap();
        let (_, helpers) = source
            .split_once(start)
            .expect("preprocessor start changed");
        let (helpers, _) = helpers.split_once(end).expect("preprocessor end changed");
        let helpers = format!("{start}{helpers}").replace(
            &format!("fn {exported}("),
            &format!("pub(super) fn {exported}("),
        );
        std::fs::write(out.join(format!("{language}_preprocessor.rs")), helpers).unwrap();
    }
}
