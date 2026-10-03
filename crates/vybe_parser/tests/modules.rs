use vybe_parser::{
    Span,
    modules::{self, Import, Module},
};

#[test]
fn local_names_exports_import_aliases_and_scoped_trivia_are_preserved() {
    let linked = modules::compile(&[
        Module { name: "outer", source: r#"WHITESPACE = _{ " " } word = { "z" } program = { SOI ~ "a" ~ imported ~ "b" ~ EOI }"#, exports: &["program"], imports: &[Import { local: "imported", module: "inner", rule: "word" }] },
        Module { name: "inner", source: r#"WHITESPACE = _{ "_" } word = { "x" ~ "y" }"#, exports: &["word"], imports: &[] },
    ]).unwrap();
    let grammar = linked.grammar();
    let tree = grammar.parse_source("outer::program", "a x_y b").unwrap();
    assert_eq!(tree.consumed(), 7);
    assert!(grammar.parse_source("outer::program", "a x y b").is_err());
    assert!(grammar.parse_source("outer::program", "a_x_y_b").is_err());
    assert!(
        grammar.rule_id("outer::word").is_none(),
        "private local rule is not exported"
    );
    assert!(grammar.recognize("inner::word", "x_y").is_ok());
    let rule = &grammar.syntax().rules[grammar.rule_id("inner::word").unwrap()];
    assert_eq!(
        linked.location(rule.name_span),
        modules::Location {
            module: 1,
            span: Span { start: 22, end: 26 }
        }
    );
}

#[test]
fn imports_of_atomic_and_nonatomic_rules_keep_mode_boundaries_and_trivia_scope() {
    let linked = modules::compile(&[
        Module {
            name: "outer",
            source: r#"WHITESPACE = _{ " " } program = @{ "a" ~ imported ~ "b" ~ EOI }"#,
            exports: &["program"],
            imports: &[Import {
                local: "imported",
                module: "inner",
                rule: "word",
            }],
        },
        Module {
            name: "inner",
            source: r#"WHITESPACE = _{ "_" } word = !{ "x" ~ "y" }"#,
            exports: &["word"],
            imports: &[],
        },
    ])
    .unwrap();
    assert!(
        linked
            .grammar()
            .recognize("outer::program", "ax_yb")
            .is_ok()
    );
    assert!(
        linked
            .grammar()
            .recognize("outer::program", "a x_y b")
            .is_err()
    );
}

#[test]
fn missing_private_colliding_duplicate_and_undefined_bindings_are_located() {
    let target = Module {
        name: "target",
        source: "private = { ANY } public = { ANY }",
        exports: &["public"],
        imports: &[],
    };
    let bad = Module {
        name: "bad",
        source: "root = { missing }",
        exports: &["root"],
        imports: &[
            Import {
                local: "root",
                module: "target",
                rule: "public",
            },
            Import {
                local: "secret",
                module: "target",
                rule: "private",
            },
        ],
    };
    let errors = modules::compile(&[bad, target]).unwrap_err();
    assert!(errors.iter().any(|error| error.code == "M003"));
    assert!(errors.iter().any(|error| error.code == "M004"));
    let undefined = errors.iter().find(|error| error.code == "G010").unwrap();
    assert_eq!(
        undefined.location,
        modules::Location {
            module: 0,
            span: Span { start: 9, end: 16 }
        }
    );
    assert_eq!(
        modules::compile(&[Module {
            name: "x",
            source: "a = { ANY } a = { ANY }",
            exports: &[],
            imports: &[]
        }])
        .unwrap_err()[0]
            .code,
        "G009"
    );
    assert_eq!(
        modules::compile(&[Module {
            name: "x",
            source: "a = { ANY }",
            exports: &["missing"],
            imports: &[]
        }])
        .unwrap_err()[0]
            .code,
        "M002"
    );
}

#[test]
fn cross_module_recursion_diagnostics_preserve_each_source_location() {
    let errors = modules::compile(&[
        Module {
            name: "a",
            source: "first = { second }",
            exports: &["first"],
            imports: &[Import {
                local: "second",
                module: "b",
                rule: "next",
            }],
        },
        Module {
            name: "b",
            source: "next = { first }",
            exports: &["next"],
            imports: &[Import {
                local: "first",
                module: "a",
                rule: "first",
            }],
        },
    ])
    .unwrap_err();
    let recursion = errors.iter().find(|error| error.code == "G011").unwrap();
    assert!(recursion.location.module < 2);
    assert!(
        recursion
            .related
            .iter()
            .any(|(location, _)| location.module != recursion.location.module)
    );
    for error in errors {
        assert!(
            error.location.span.end
                <= ["first = { second }", "next = { first }"][error.location.module].len()
        );
    }
}

#[test]
fn nullable_module_trivia_remains_a_compile_error_and_stack_effects_cross_imports() {
    let bad = modules::compile(&[Module {
        name: "a",
        source: r#"WHITESPACE = _{ "" } root = { "x" ~ "y" }"#,
        exports: &["root"],
        imports: &[],
    }])
    .unwrap_err();
    assert!(bad.iter().any(|error| error.code == "G012"));
    let linked = modules::compile(&[
        Module {
            name: "a",
            source: "root = { imported ~ POP ~ EOI }",
            exports: &["root"],
            imports: &[Import {
                local: "imported",
                module: "b",
                rule: "push",
            }],
        },
        Module {
            name: "b",
            source: r#"push = { PUSH("x") }"#,
            exports: &["push"],
            imports: &[],
        },
    ])
    .unwrap();
    assert!(linked.grammar().recognize("a::root", "xx").is_ok());
    let root = linked.grammar().rule_id("a::root").unwrap();
    assert!(
        linked.grammar().analysis().rules[root]
            .effects
            .is_stack_dependent()
    );
}
