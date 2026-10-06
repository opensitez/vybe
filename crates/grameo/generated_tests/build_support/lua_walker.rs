use std::{fs, path::Path};

pub fn generate(out: &Path) {
    // Compile the actual language walker against owned native pairs. The
    // normalization pass depends on the compiler and remains outside this
    // isolated raw-AST migration check; it is not replaced or stubbed in the
    // production frontend.
    let walker_path = "../../../languages/lua/src/walker.rs";
    println!("cargo:rerun-if-changed={walker_path}");
    let walker = fs::read_to_string(walker_path).unwrap();
    let marker = "    super::normalize::normalize_module(&mut module);";
    assert_eq!(walker.matches(marker).count(), 1, "Lua normalization marker changed");
    let native_walker = walker
        .replace(marker, "")
        .replace("    let mut module = Module {", "    let module = Module {");
    // Pest is confined to the independent reference test harness.
    let raw_walker = native_walker.replace(
        "use grameo::compat::Pair;",
        "use pest::Parser;\nuse pest::iterators::Pair;",
    );
    let expression_adapter = "\n#[allow(dead_code)]\npub fn expression_for_conformance(source: &str) -> Result<(Expression, usize), String> {\n    let _index = vybe_ast::line_index::LineIndex::install(source);\n    let pair = super::LuaParser::parse(super::Rule::expr, source).map_err(|e| e.to_string())?.next().ok_or(\"missing expression\")?;\n    let end = pair.as_span().end();\n    Ok((walk_expression(pair)?, end))\n}\n";
    fs::write(
        out.join("lua_native_walker.rs"),
        native_walker + expression_adapter,
    )
    .unwrap();
    fs::write(
        out.join("lua_raw_walker.rs"),
        raw_walker + expression_adapter,
    )
    .unwrap();
    let (_, literal_helpers) = walker
        .split_once("fn unescape_lua_string(")
        .expect("Lua literal helper changed");
    let (literal_helpers, _) = literal_helpers
        .split_once("fn walk_primary(")
        .expect("Lua literal helper boundary changed");
    fs::write(
        out.join("lua_literal_helpers.rs"),
        format!("pub(super) fn unescape_lua_string({literal_helpers}").replace(
            "fn strip_long_brackets(",
            "pub(super) fn strip_long_brackets(",
        ),
    )
    .unwrap();
    let normalize_path = "../../../languages/lua/src/normalize.rs";
    println!("cargo:rerun-if-changed={normalize_path}");
    let normalization = fs::read_to_string(normalize_path).unwrap();
    let (_, numeric_helpers) = normalization
        .split_once("fn lua_expr_contains_unshadowed_ident(")
        .expect("Lua numeric helper boundary changed");
    let (numeric_predicates, _) = numeric_helpers
        .split_once("fn lua_replace_unshadowed_ident(")
        .expect("Lua numeric predicate boundary changed");
    fs::write(
        out.join("lua_numeric_predicates.rs"),
        format!("use vybe_ast::*;\nfn lua_expr_contains_unshadowed_ident({numeric_predicates}")
            .replace(
                "fn lua_expr_contains_unshadowed_ident(",
                "pub(super) fn lua_expr_contains_unshadowed_ident(",
            ),
    )
    .unwrap();
    let (_, ident) = normalization
        .split_once("fn lua_ident(")
        .expect("Lua identifier helper changed");
    let (ident, _) = ident
        .split_once("fn lua_optional_param(")
        .expect("Lua identifier helper boundary changed");
    fs::write(out.join("lua_numeric_helpers.rs"), format!("use vybe_ast::*;\nfn lua_ident({ident}\nfn lua_expr_contains_unshadowed_ident({numeric_helpers}")).unwrap();
}
