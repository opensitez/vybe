use std::{env, fs, path::PathBuf};

fn main() {
    let out = PathBuf::from(env::var_os("OUT_DIR").unwrap());
    println!("cargo:rerun-if-changed=tests/modular_outer.grammar");
    println!("cargo:rerun-if-changed=tests/modular_inner.grammar");
    let outer = fs::read_to_string("tests/modular_outer.grammar").unwrap();
    let inner = fs::read_to_string("tests/modular_inner.grammar").unwrap();
    let generated = vybe_parser_codegen::generate_modules(&[
        vybe_parser_codegen::Module {
            name: "outer",
            source: &outer,
            exports: &["program"],
            imports: &[vybe_parser_codegen::Import {
                local: "imported",
                module: "inner",
                rule: "word",
            }],
        },
        vybe_parser_codegen::Module {
            name: "inner",
            source: &inner,
            exports: &["word"],
            imports: &[],
        },
    ])
    .unwrap();
    fs::write(out.join("modular.rs"), generated).unwrap();
    for (name, path) in [
        ("fixtures", "../conformance_tests/tests/fixtures.pest"),
        ("lua", "../../../languages/lua/src/grammar.pest"),
        ("ast", "tests/ast.grammar"),
        ("expression", "tests/expression.grammar"),
    ] {
        println!("cargo:rerun-if-changed={path}");
        let source = fs::read_to_string(path).unwrap();
        let generated = vybe_parser_codegen::generate(&source)
            .unwrap_or_else(|errors| panic!("{path}: {errors:?}"));
        fs::write(out.join(format!("{name}.rs")), generated).unwrap();
    }
    let mut modules = String::new();
    if env::var_os("CARGO_FEATURE_ALL_GRAMMARS").is_some() {
        let root = PathBuf::from("../../../languages");
        println!("cargo:rerun-if-changed={}", root.display());
        let mut paths: Vec<_> = fs::read_dir(root)
            .unwrap()
            .map(|entry| entry.unwrap().path())
            .collect();
        paths.sort();
        for path in paths {
            let grammar = path.join("src/grammar.pest");
            if !grammar.exists() {
                continue;
            }
            let name = path.file_name().unwrap().to_str().unwrap();
            println!("cargo:rerun-if-changed={}", grammar.display());
            let source = fs::read_to_string(&grammar).unwrap();
            let generated = vybe_parser_codegen::generate(&source).unwrap();
            let output = out.join(format!("all_{name}.rs"));
            fs::write(&output, generated).unwrap();
            modules.push_str(&format!("pub mod {name} {{ include!({:?}); }}\n", output));
        }
    }
    fs::write(out.join("all_grammars.rs"), modules).unwrap();
}
