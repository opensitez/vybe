//! Compile the production JavaScript AST walk in isolation from runtime plugins.
//! The only compiler dependency is two naming functions, copied verbatim.
use std::{fs, path::Path};
pub fn generate(out: &Path) {
    for path in [
        "../../../languages/js/src/walker.rs",
        "../../../languages/js/src/lib.rs",
        "../../vybe_compiler/src/primitives/class_slots.rs",
    ] {
        println!("cargo:rerun-if-changed={path}");
    }
    let walker = fs::read_to_string("../../../languages/js/src/walker.rs")
        .unwrap()
        .replace(
            "use vybe_compiler::primitives::class_slots;",
            "use crate::class_slots;",
        );
    fs::write(out.join("javascript_reference_walker.rs"), &walker).unwrap();
    let previous_walker = walker
        .replace("let presents_ui = references_document(&mut body);", "let presents_ui = references_document(&body);")
        .replace("fn references_document(body: &mut [Statement])", "fn references_document(body: &[Statement])")
        .replace("    for stmt in body {\n        stmt.walk_exprs_mut", "    for stmt in body {\n        let mut stmt = stmt.clone();\n        stmt.walk_exprs_mut");
    let hoist_start = previous_walker
        .find("    // Most small scripts have no declarations to move.")
        .unwrap();
    let hoist_end = previous_walker[hoist_start..]
        .find("\n    // Const-folding pass")
        .unwrap()
        + hoist_start;
    let mut previous_walker = previous_walker;
    previous_walker.replace_range(
        hoist_start..hoist_end,
        r#"    let mut hoisted = Vec::new();
    let mut rest = Vec::new();
    for stmt in body {
        if matches!(stmt.kind, StmtKind::FunctionDecl { .. }) {
            hoisted.push(stmt);
        } else {
            rest.push(stmt);
        }
    }
    hoisted.append(&mut rest);
    let mut body = hoisted;
"#,
    );
    fs::write(out.join("javascript_previous_walker.rs"), previous_walker).unwrap();
    let native = walker.replace("use pest::Parser;", "").replace(
        "use pest::iterators::Pair;",
        "use grameo::compat::Pair;",
    );
    fs::write(out.join("javascript_native_walker.rs"), native).unwrap();
    let slots = fs::read_to_string("../../vybe_compiler/src/primitives/class_slots.rs").unwrap();
    let helpers = ["getter_name", "setter_name"]
        .map(|name| {
            let start = slots.find(&format!("pub fn {name}(")).unwrap();
            let end = slots[start..].find("\n}").unwrap() + start + 2;
            &slots[start..end]
        })
        .join("\n");
    fs::write(out.join("javascript_class_slots.rs"), helpers).unwrap();
    let library = fs::read_to_string("../../../languages/js/src/lib.rs").unwrap();
    let start = library.find("const JS_PRELUDE:").unwrap();
    let end = library.find("/// Embedded profile TOML source.").unwrap();
    let frontend = &library[start..end];
    fs::write(out.join("javascript_frontend.rs"), frontend).unwrap();
    // Previous allocation strategy; preserve the same production prelude,
    // directive handling, body boundary marker, and walker as the current copy.
    // Baseline is the immediately preceding frontend: borrowed prelude cache,
    // cloned UI detection, and the missing strict directive capacity slot.
    let previous = frontend.replace(" + usize::from(user_is_strict)", "");
    fs::write(out.join("javascript_previous_frontend.rs"), previous).unwrap();
}
