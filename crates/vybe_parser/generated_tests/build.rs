#[path = "build_support/lua_walker.rs"]
mod lua_walker;
use std::{env, fs, path::PathBuf};

fn main() {
    let out = PathBuf::from(env::var_os("OUT_DIR").unwrap());
    lua_walker::generate(&out);
    println!("cargo:rerun-if-changed=../tests/islands.grammar");
    let source = fs::read_to_string("../tests/islands.grammar").unwrap();
    let mut grammar = vybe_parser::compile(&source).unwrap();
    use vybe_parser::{
        islands::{Operator, TrailingTrivia},
        pratt::Fixity,
    };
    grammar
        .bind_pratt(
            "expr",
            "atom",
            &[
                Operator {
                    rule: "minus",
                    precedence: 25,
                    fixity: Fixity::Prefix,
                },
                Operator {
                    rule: "power",
                    precedence: 30,
                    fixity: Fixity::InfixRight,
                },
                Operator {
                    rule: "bang",
                    precedence: 40,
                    fixity: Fixity::Postfix,
                },
                Operator {
                    rule: "star",
                    precedence: 20,
                    fixity: Fixity::InfixLeft,
                },
                Operator {
                    rule: "plus",
                    precedence: 10,
                    fixity: Fixity::InfixLeft,
                },
                Operator {
                    rule: "minus",
                    precedence: 10,
                    fixity: Fixity::InfixLeft,
                },
            ],
            TrailingTrivia::Consume,
        )
        .unwrap();
    fs::write(
        out.join("islands.rs"),
        vybe_parser_codegen::emit(&grammar).unwrap(),
    )
    .unwrap();
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
        if name == "lua" {
            let mut grammar = vybe_parser::compile(&source).unwrap();
            let mut operators = vec![Operator {
                rule: "unop",
                precedence: 25,
                fixity: Fixity::Prefix,
            }];
            for (rule, precedence, fixity) in [
                ("pow_op", 30, Fixity::InfixRight),
                ("mul_op", 20, Fixity::InfixLeft),
                ("additive_op", 19, Fixity::InfixLeft),
                ("CONCAT", 18, Fixity::InfixRight),
                ("shift_op", 17, Fixity::InfixLeft),
                ("AMP", 16, Fixity::InfixLeft),
                ("TILDE", 15, Fixity::InfixLeft),
                ("PIPE", 14, Fixity::InfixLeft),
                ("compare_op", 13, Fixity::InfixLeft),
                ("KW_AND", 12, Fixity::InfixLeft),
                ("KW_OR", 11, Fixity::InfixLeft),
            ] {
                operators.push(Operator {
                    rule,
                    precedence,
                    fixity,
                });
            }
            grammar
                .bind_pratt("expr", "postfix", &operators, TrailingTrivia::Consume)
                .unwrap();
            fs::write(
                out.join("lua_pratt.rs"),
                vybe_parser_codegen::emit(&grammar).unwrap(),
            )
            .unwrap();
        }
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
