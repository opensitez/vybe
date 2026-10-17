#[path = "build_support/lua_walker.rs"]
mod lua_walker;
use std::{env, fs, path::PathBuf};

fn main() {
    let out = PathBuf::from(env::var_os("OUT_DIR").unwrap());
    println!("cargo:rerun-if-changed=tests/common_ast.grammar");
    let ast_source = fs::read_to_string("tests/common_ast.grammar").unwrap();
    fs::write(out.join("common_ast.rs"), grameo_codegen::generate_with_ast_schema(&ast_source, vybe_ast_bindings::SCHEMA).unwrap()).unwrap();
    println!("cargo:rerun-if-changed=tests/common_integer.grammar");
    let source = fs::read_to_string("tests/common_integer.grammar").unwrap();
    fs::write(out.join("common_integer.rs"), grameo_codegen::generate_with_ast_schema(&source, vybe_ast_bindings::SCHEMA).unwrap()).unwrap();
    println!("cargo:rerun-if-changed=tests/normal_class.grammar");
    let source = fs::read_to_string("tests/normal_class.grammar").unwrap();
    fs::write(out.join("normal_class.rs"), grameo_codegen::generate_with_ast_schema(&source, vybe_ast_bindings::SCHEMA).unwrap()).unwrap();
    let api_grammar = "root = { token ~ EOI } token = @{ ASCII_ALPHA+ }";
    for (name, api) in [
        ("capture_api", grameo_codegen::ParserApi::CAPTURE_ONLY),
        ("hybrid_api", grameo_codegen::ParserApi::HYBRID_CAPTURES),
    ] {
        fs::write(out.join(format!("{name}.rs")), grameo_codegen::generate_with_api(api_grammar, api).unwrap()).unwrap();
    }
    lua_walker::generate(&out);
    println!("cargo:rerun-if-changed=../tests/islands.grammar");
    let source = fs::read_to_string("../tests/islands.grammar").unwrap();
    let mut grammar = grameo::compile(&source).unwrap();
    use grameo::{
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
        grameo_codegen::emit(&grammar).unwrap(),
    )
    .unwrap();
    println!("cargo:rerun-if-changed=tests/modular_outer.grammar");
    println!("cargo:rerun-if-changed=tests/modular_inner.grammar");
    let outer = fs::read_to_string("tests/modular_outer.grammar").unwrap();
    let inner = fs::read_to_string("tests/modular_inner.grammar").unwrap();
    let generated = grameo_codegen::generate_modules(&[
        grameo_codegen::Module {
            name: "outer",
            source: &outer,
            exports: &["program"],
            imports: &[grameo_codegen::Import {
                local: "imported",
                module: "inner",
                rule: "word",
            }],
        },
        grameo_codegen::Module {
            name: "inner",
            source: &inner,
            exports: &["word"],
            imports: &[],
        },
    ])
    .unwrap();
    fs::write(out.join("modular.rs"), generated).unwrap();
    println!("cargo:rerun-if-changed=../../../languages/php/src/grammar.pest");
    let php = grameo_codegen::generate(&fs::read_to_string("../../../languages/php/src/grammar.pest").unwrap()).unwrap();
    fs::write(out.join("php.rs"), &php).unwrap();
    let marker = "    fn delimited_fragment(";
    let start = php.find(marker).unwrap();
    let end = start + php[start..].find("    }}\n").unwrap() + "    }}\n".len();
    let mut delimited_baseline = php.clone();
    delimited_baseline.replace_range(start..end,
        "    fn delimited_fragment(&self, _:usize)->Option<::grameo::quoted::DelimitedFragment<'_>>{ None }\n");
    fs::write(out.join("php_delimited_baseline.rs"), delimited_baseline).unwrap();
    // Preserve every other generated optimization when measuring semantic
    // trivia/lookahead dispatch against the original scheduled matcher.
    let semantic_baseline = php.replace(
        "fn atomic_expression(&self, expression: usize) -> bool { ATOMIC_EXPRESSIONS[expression] }",
        "fn atomic_expression(&self, _: usize) -> bool { false }");
    assert_ne!(semantic_baseline, php);
    let semantic_baseline = semantic_baseline.lines()
        .filter(|line| !line.starts_with("static ATOMIC_EXPRESSIONS:"))
        .collect::<Vec<_>>().join("\n") + "\n";
    fs::write(out.join("php_semantic_baseline.rs"), semantic_baseline).unwrap();
    // Paired native-program oracle: all generated tables and dispatch stay
    // identical, with only the quoted-token certificate disabled.
    let marker = "    fn quoted_fragment(";
    let start = php.find(marker).expect("PHP quoted certificate");
    let end = start + php[start..].find("    }}\n").unwrap() + "    }}\n".len();
    let mut baseline = php.clone();
    baseline.replace_range(start..end,
        "    fn quoted_fragment(&self, _:usize)->Option<::grameo::quoted::Fragment>{ None }\n");
    fs::write(out.join("php_baseline.rs"), baseline).unwrap();
    for (name, path) in [
        ("fixtures", "../conformance_tests/tests/fixtures.pest"),
        ("terminal_index", "tests/terminal_index.grammar"),
        ("lua", "../../../languages/lua/src/grammar.pest"),
        ("ast", "tests/ast.grammar"),
        ("expression", "tests/expression.grammar"),
        ("quoted", "tests/quoted.grammar"),
    ] {
        println!("cargo:rerun-if-changed={path}");
        let source = fs::read_to_string(path).unwrap();
        let generated = grameo_codegen::generate(&source)
            .unwrap_or_else(|errors| panic!("{path}: {errors:?}"));
        fs::write(out.join(format!("{name}.rs")), generated).unwrap();
        if name == "lua" {
            let mut grammar = grameo::compile(&source).unwrap();
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
            grammar.set_literal_choice_failures(true);
            fs::write(
                out.join("lua_pratt.rs"),
                grameo_codegen::emit(&grammar).unwrap(),
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
            let generated = grameo_codegen::generate(&source).unwrap();
            let output = out.join(format!("all_{name}.rs"));
            fs::write(&output, generated).unwrap();
            modules.push_str(&format!("pub mod {name} {{ include!({:?}); }}\n", output));
        }
    }
    fs::write(out.join("all_grammars.rs"), modules).unwrap();
}
