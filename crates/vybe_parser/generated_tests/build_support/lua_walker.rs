use std::{fs, path::Path};

pub fn generate(out: &Path) {
    // Compile the actual language walker against owned native pairs. The
    // normalization pass depends on the compiler and remains outside this
    // isolated raw-AST migration check; it is not replaced or stubbed in the
    // production frontend.
    let walker_path = "../../../languages/lua/src/walker.rs";
    println!("cargo:rerun-if-changed={walker_path}");
    let walker = fs::read_to_string(walker_path).unwrap();
    let walker = walker
        .replace("#[cfg(not(feature = \"native-parser\"))]\n", "")
        .replace(
            "#[cfg(feature = \"native-parser\")]\nuse vybe_parser::compat::Pair;\n",
            "",
        );
    for marker in [
        "use pest::Parser;",
        "use pest::iterators::Pair;",
        "    super::normalize::normalize_module(&mut module);",
    ] {
        assert_eq!(
            walker.matches(marker).count(),
            1,
            "Lua walker harness marker changed: {marker}"
        );
    }
    let raw_walker = walker
        .replace("    super::normalize::normalize_module(&mut module);", "")
        .replace("    let mut module = Module {", "    let module = Module {");
    let native_walker = raw_walker.replace("use pest::Parser;", "").replace(
        "use pest::iterators::Pair;",
        "use vybe_parser::compat::Pair;",
    );
    fs::write(out.join("lua_native_walker.rs"), native_walker).unwrap();
    fs::write(out.join("lua_raw_walker.rs"), raw_walker).unwrap();
    let normalize_path = "../../../languages/lua/src/normalize.rs";
    println!("cargo:rerun-if-changed={normalize_path}");
    let normalization = fs::read_to_string(normalize_path).unwrap();
    let (_, numeric_helpers) = normalization
        .split_once("fn lua_expr_contains_unshadowed_ident(")
        .expect("Lua numeric helper boundary changed");
    let (_, ident) = normalization
        .split_once("fn lua_ident(")
        .expect("Lua identifier helper changed");
    let (ident, _) = ident
        .split_once("fn lua_optional_param(")
        .expect("Lua identifier helper boundary changed");
    fs::write(out.join("lua_numeric_helpers.rs"), format!("use vybe_ast::*;\nfn lua_ident({ident}\nfn lua_expr_contains_unshadowed_ident({numeric_helpers}")).unwrap();
}
