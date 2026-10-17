use grameo::{
    Span,
    ast::builder::{Argument, AstBuilder, Target},
    ast::{Binding, Cardinality, ConstructorSchema, Field, FieldKind, FieldSchema, Source},
    builder::{BuildError, Builder},
};

const GRAMMAR: &str = r#"
// @grameo ast root = Block(items: statement*, span: rule)
root = { statement* ~ EOI }
// @grameo ast statement = Return(value: integer?, span: rule)
statement = { "return" ~ integer? ~ ";" }
// @grameo ast integer = Integer(text: source, span: rule)
integer = @{ ASCII_DIGIT+ }
WHITESPACE = _{ " " | "\n" }
"#;
static BINDINGS: &[Binding] = &[
    Binding {
        rule: 0,
        constructor: 0,
        fields: &[
            Field {
                name: "items",
                source: Source::Child {
                    rule: 1,
                    occurrence: None,
                    cardinality: Cardinality::Many,
                },
            },
            Field {
                name: "span",
                source: Source::Span,
            },
        ],
    },
    Binding {
        rule: 1,
        constructor: 2,
        fields: &[
            Field {
                name: "value",
                source: Source::Child {
                    rule: 2,
                    occurrence: None,
                    cardinality: Cardinality::Optional,
                },
            },
            Field {
                name: "span",
                source: Source::Span,
            },
        ],
    },
    Binding {
        rule: 2,
        constructor: 1,
        fields: &[
            Field {
                name: "text",
                source: Source::Text,
            },
            Field {
                name: "span",
                source: Source::Span,
            },
        ],
    },
];
#[derive(Debug, PartialEq)]
enum Node {
    Integer(i64, Span),
    Return(Option<Box<Node>>, Span),
    Block(Vec<Node>, Span),
}
struct Nodes;
impl Target for Nodes {
    type Value = Node;
    type Undo = ();
    type Context = ();
    fn prepare(_: &str) {}
    fn validate(
        _: &(),
        constructor: usize,
        _: &[Field],
        span: Span,
        text: &str,
    ) -> Result<(), BuildError> {
        if constructor == 1 && text.parse::<i64>().is_err() {
            Err(BuildError {
                span,
                message: "integer exceeds target range".into(),
            })
        } else {
            Ok(())
        }
    }
    fn construct(
        _: &(),
        constructor: usize,
        span: Span,
        args: Vec<Argument<'_, Node>>,
    ) -> (Node, ()) {
        let first = args.into_iter().next().unwrap();
        (
            match (constructor, first) {
                (0, Argument::Many(items)) => Node::Block(items, span),
                (1, Argument::Text(text)) => Node::Integer(text.parse().unwrap(), span),
                (2, Argument::Optional(value)) => Node::Return(value.map(Box::new), span),
                _ => panic!("invalid target constructor signature"),
            },
            (),
        )
    }
    fn split(_: usize, value: Node, _: ()) -> Vec<Node> {
        match value {
            Node::Integer(..) => vec![],
            Node::Return(value, _) => value.into_iter().map(|v| *v).collect(),
            Node::Block(items, _) => items,
        }
    }
}

#[test]
fn annotations_build_owned_statements_without_capture_walks() {
    let grammar = grameo::compile_with_ast_bindings(GRAMMAR).unwrap();
    let source = "return 7;\nreturn;";
    let mut builder = AstBuilder::<Nodes>::new(BINDINGS, source);
    let consumed = grameo::engine::build_ast_program(
        &grammar,
        "root",
        source,
        Default::default(),
        &mut builder,
    )
    .unwrap();
    assert_eq!(consumed, source.len());
    assert_eq!(
        builder.into_value().unwrap(),
        Node::Block(
            vec![
                Node::Return(
                    Some(Box::new(Node::Integer(7, Span { start: 7, end: 8 }))),
                    Span { start: 0, end: 9 }
                ),
                Node::Return(None, Span { start: 10, end: 17 }),
            ],
            Span { start: 0, end: 17 }
        )
    );
}

#[test]
fn speculative_constructor_values_rollback_and_repeated_marks_restore() {
    let source = "return 7;";
    let grammar = grameo::compile_with_ast_bindings(&GRAMMAR.replace(
        "statement* ~ EOI",
        "((statement ~ \"x\") | statement) ~ EOI",
    ))
    .unwrap();
    let mut builder = AstBuilder::<Nodes>::new(BINDINGS, source);
    let mark = builder.checkpoint();
    grameo::engine::build_ast_program(&grammar, "root", source, Default::default(), &mut builder)
        .unwrap();
    assert!(matches!(builder.into_value().unwrap(), Node::Block(items, _) if items.len() == 1));
    let mut builder = AstBuilder::<Nodes>::new(BINDINGS, source);
    builder.begin(2, 0);
    builder
        .try_finish(2, Span { start: 0, end: 1 }, "7")
        .unwrap();
    builder.rollback(mark);
    builder.rollback(mark);
    assert!(builder.into_value().is_err());
}

#[test]
fn target_validation_is_fatal_and_restores_the_builder() {
    let source = "return 99999999999999999999999999999999;";
    let grammar = grameo::compile_with_ast_bindings(GRAMMAR).unwrap();
    let mut builder = AstBuilder::<Nodes>::new(BINDINGS, source);
    let mark = builder.checkpoint();
    let error = grameo::engine::build_ast_program(
        &grammar,
        "root",
        source,
        Default::default(),
        &mut builder,
    )
    .unwrap_err();
    assert_eq!(error.kind, grameo::ParseErrorKind::BuildFailure);
    assert_eq!(builder.checkpoint(), mark);
}

#[test]
fn metadata_is_only_read_from_real_line_comments_and_diagnostics_are_located() {
    assert!(
        grameo::compile_with_ast_bindings(r#"root = { "// @grameo ast fake = Fake()" }"#)
            .unwrap()
            .syntax()
            .ast_bindings
            .is_empty()
    );
    assert!(
        grameo::compile_with_ast_bindings("/* // @grameo ast fake = Fake() */ root = { \"x\" }")
            .unwrap()
            .syntax()
            .ast_bindings
            .is_empty()
    );
    for source in [
        "// @grameo ast missing = Value()\nroot = { \"x\" }",
        "// @grameo ast root = Value(a: missing)\nroot = { \"x\" }",
    ] {
        let errors = grameo::compile_with_ast_bindings(source).unwrap_err();
        assert_eq!(errors[0].code, "G041");
        assert_eq!(errors[0].span.start, 0);
    }
    assert_eq!(
        grameo::compile_with_ast_bindings("// @grameo ast root = Value(a integer)\nroot = { \"x\" }").unwrap_err()[0]
            .code,
        "G040"
    );
    assert_eq!(
        grameo::compile_with_ast_bindings("// @grameo ast\nroot = { \"x\" }").unwrap_err()[0].code,
        "G040"
    );
}

#[test]
fn linked_modules_preserve_annotation_rule_identity_and_locations() {
    let modules = [grameo::modules::Module {
        name: "owned",
        source: GRAMMAR,
        exports: &["root"],
        imports: &[],
    }];
    let linked = grameo::modules::compile_with_ast_bindings(&modules).unwrap();
    let bindings = &linked.grammar().syntax().ast_bindings;
    assert_eq!(bindings.len(), 3);
    for binding in bindings {
        assert!(linked.grammar().rule_id(&binding.rule).is_some());
    }
    let grameo::ast::SourceSyntax::Child { rule, .. } = &bindings[0].fields[0].source else {
        panic!()
    };
    assert_eq!(rule, &bindings[1].rule);
}

#[test]
fn target_schema_checks_child_types_and_optional_list_cardinality() {
    let grammar = grameo::compile_with_ast_bindings(GRAMMAR).unwrap();
    let schema = [
        ConstructorSchema {
            name: "Integer",
            value_kind: "expression",
            fields: &[
                FieldSchema {
                    name: "text",
                    kind: FieldKind::Text,
                },
                FieldSchema {
                    name: "span",
                    kind: FieldKind::Span,
                },
            ],
        },
        ConstructorSchema {
            name: "Return",
            value_kind: "statement",
            fields: &[
                FieldSchema {
                    name: "value",
                    kind: FieldKind::Child {
                        value_kind: "expression",
                        cardinality: Cardinality::Optional,
                    },
                },
                FieldSchema {
                    name: "span",
                    kind: FieldKind::Span,
                },
            ],
        },
        ConstructorSchema {
            name: "Block",
            value_kind: "block",
            fields: &[
                FieldSchema {
                    name: "items",
                    kind: FieldKind::Child {
                        value_kind: "statement",
                        cardinality: Cardinality::Many,
                    },
                },
                FieldSchema {
                    name: "span",
                    kind: FieldKind::Span,
                },
            ],
        },
    ];
    grammar.validate_ast_schema(&schema).unwrap();
    let incorrect =
        grameo::compile_with_ast_bindings(&GRAMMAR.replace("value: integer?", "value: integer*")).unwrap();
    assert_eq!(
        incorrect.validate_ast_schema(&schema).unwrap_err()[0].code,
        "G042"
    );
}

#[test]
fn all_children_selector_checks_types_across_alternatives_and_unbound_wrappers() {
    let source = r#"
// @grameo ast root = Statement(value: children)
root = { wrapper ~ EOI }
wrapper = _{ first | second }
// @grameo ast first = Return()
first = { "a" }
// @grameo ast second = Return()
second = { "b" }
"#;
    let fields = [FieldSchema {
        name: "value",
        kind: FieldKind::Child {
            value_kind: "statement",
            cardinality: Cardinality::One,
        },
    }];
    let schema = [
        ConstructorSchema {
            name: "Statement",
            value_kind: "statement",
            fields: &fields,
        },
        ConstructorSchema {
            name: "Return",
            value_kind: "statement",
            fields: &[],
        },
        ConstructorSchema {
            name: "Integer",
            value_kind: "expression",
            fields: &[],
        },
    ];
    grameo::compile_with_ast_bindings(source)
        .unwrap()
        .validate_ast_schema(&schema)
        .unwrap();
    let mixed =
        grameo::compile_with_ast_bindings(&source.replace("second = Return()", "second = Integer()")).unwrap();
    assert_eq!(
        mixed.validate_ast_schema(&schema).unwrap_err()[0].code,
        "G042"
    );
}

#[test]
fn default_peg_ignores_ast_annotation_comments() {
    for annotation in ["// @grameo ast", "// @grameo ast missing = Unknown()", "// @grameo ast root = Value(a integer)"] {
        let source = format!("{annotation}\nroot = {{ \"x\" }}");
        let grammar = grameo::compile(&source).unwrap();
        assert!(grammar.syntax().ast_bindings.is_empty());
        assert!(grammar.parse_source("root", "x").is_ok());
        assert!(grameo::compile_with_ast_bindings(&source).is_err());
    }
}
