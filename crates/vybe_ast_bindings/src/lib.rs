//! Shared language-neutral AST constructors for grammar-driven frontends.
//! Grameo itself and the common AST crate remain independent of this bridge.
use grameo::{
    Span as Bytes,
    ast::builder::{Argument, Target},
};
use vybe_ast::class_normalize::{
    Access, NormalClass, NormalField, NormalMembers, NormalMethod, build_normal_method,
};
use vybe_ast::{
    ClassMember, ExprKind, Expression, Lang, Literal, Module, ScopeDeclKind, Span, Statement,
    StmtKind,
};
#[derive(Debug, Clone, Copy)]
#[repr(usize)]
pub enum ConstructorId {
    Block,
    Class,
    Function,
    Integer,
    Module,
    Name,
    Return,
    Statement,
    While,
    Namespace,
    NormalMethod,
    NormalStaticMethod,
    NormalField,
    NormalStaticField,
    NormalClass,
    NormalMember,
    ScopeClosed,
    ScopeGlobal,
    ScopeNonlocal,
    ProtocolSlot,
    ProtocolFunction,
    DirectiveBlock,
    DirectiveModule,
}
use ConstructorId as C;
use grameo::ast::{
    Cardinality as Card, ConstructorSchema as Constructor, FieldKind as Kind, FieldSchema as Field,
};

/// Generate using this schema so constructor IDs match `ConstructorId`.
pub static SCHEMA: &[Constructor<'static>] = &[
    Constructor {
        name: "Block",
        value_kind: "body",
        fields: &[Field {
            name: "body",
            kind: Kind::Child {
                value_kind: "statement",
                cardinality: Card::Many,
            },
        }],
    },
    Constructor {
        name: "Class",
        value_kind: "statement",
        fields: &[
            Field {
                name: "name",
                kind: Kind::Child {
                    value_kind: "name",
                    cardinality: Card::One,
                },
            },
            Field {
                name: "methods",
                kind: Kind::Child {
                    value_kind: "statement",
                    cardinality: Card::Many,
                },
            },
            Field {
                name: "span",
                kind: Kind::Span,
            },
        ],
    },
    Constructor {
        name: "Function",
        value_kind: "statement",
        fields: &[
            Field {
                name: "name",
                kind: Kind::Child {
                    value_kind: "name",
                    cardinality: Card::One,
                },
            },
            Field {
                name: "body",
                kind: Kind::Child {
                    value_kind: "body",
                    cardinality: Card::One,
                },
            },
            Field {
                name: "span",
                kind: Kind::Span,
            },
        ],
    },
    Constructor {
        name: "Integer",
        value_kind: "expression",
        fields: &[
            Field {
                name: "text",
                kind: Kind::Text,
            },
            Field {
                name: "span",
                kind: Kind::Span,
            },
        ],
    },
    Constructor {
        name: "Module",
        value_kind: "module",
        fields: &[Field {
            name: "body",
            kind: Kind::Child {
                value_kind: "statement",
                cardinality: Card::Many,
            },
        }],
    },
    Constructor {
        name: "Name",
        value_kind: "name",
        fields: &[Field {
            name: "text",
            kind: Kind::Text,
        }],
    },
    Constructor {
        name: "Return",
        value_kind: "statement",
        fields: &[
            Field {
                name: "value",
                kind: Kind::Child {
                    value_kind: "expression",
                    cardinality: Card::Optional,
                },
            },
            Field {
                name: "span",
                kind: Kind::Span,
            },
        ],
    },
    Constructor {
        name: "Statement",
        value_kind: "statement",
        fields: &[Field {
            name: "value",
            kind: Kind::Child {
                value_kind: "statement",
                cardinality: Card::One,
            },
        }],
    },
    Constructor {
        name: "While",
        value_kind: "statement",
        fields: &[
            Field {
                name: "condition",
                kind: Kind::Child {
                    value_kind: "expression",
                    cardinality: Card::One,
                },
            },
            Field {
                name: "body",
                kind: Kind::Child {
                    value_kind: "body",
                    cardinality: Card::One,
                },
            },
            Field {
                name: "span",
                kind: Kind::Span,
            },
        ],
    },
    Constructor {
        name: "Namespace",
        value_kind: "statement",
        fields: &[
            Field {
                name: "name",
                kind: Kind::Child {
                    value_kind: "name",
                    cardinality: Card::One,
                },
            },
            Field {
                name: "body",
                kind: Kind::Child {
                    value_kind: "body",
                    cardinality: Card::One,
                },
            },
            Field {
                name: "span",
                kind: Kind::Span,
            },
        ],
    },
    Constructor {
        name: "NormalMethod",
        value_kind: "normal_member",
        fields: METHOD_FIELDS,
    },
    Constructor {
        name: "NormalStaticMethod",
        value_kind: "normal_member",
        fields: METHOD_FIELDS,
    },
    Constructor {
        name: "NormalField",
        value_kind: "normal_member",
        fields: FIELD_FIELDS,
    },
    Constructor {
        name: "NormalStaticField",
        value_kind: "normal_member",
        fields: FIELD_FIELDS,
    },
    Constructor {
        name: "NormalClass",
        value_kind: "normal_class",
        fields: &[
            Field {
                name: "name",
                kind: Kind::Child {
                    value_kind: "name",
                    cardinality: Card::One,
                },
            },
            Field {
                name: "members",
                kind: Kind::Child {
                    value_kind: "normal_member",
                    cardinality: Card::Many,
                },
            },
            Field {
                name: "span",
                kind: Kind::Span,
            },
        ],
    },
    Constructor {
        name: "NormalMember",
        value_kind: "normal_member",
        fields: &[Field {
            name: "value",
            kind: Kind::Child {
                value_kind: "normal_member",
                cardinality: Card::One,
            },
        }],
    },
    Constructor {
        name: "ScopeClosed",
        value_kind: "statement",
        fields: SCOPE_FIELDS,
    },
    Constructor {
        name: "ScopeGlobal",
        value_kind: "statement",
        fields: SCOPE_FIELDS,
    },
    Constructor {
        name: "ScopeNonlocal",
        value_kind: "statement",
        fields: SCOPE_FIELDS,
    },
    Constructor {
        name: "ProtocolSlot",
        value_kind: "protocol_slot",
        fields: &[Field {
            name: "text",
            kind: Kind::Text,
        }],
    },
    Constructor {
        name: "ProtocolFunction",
        value_kind: "statement",
        fields: &[
            Field {
                name: "name",
                kind: Kind::Child {
                    value_kind: "name",
                    cardinality: Card::One,
                },
            },
            Field {
                name: "body",
                kind: Kind::Child {
                    value_kind: "body",
                    cardinality: Card::One,
                },
            },
            Field {
                name: "slot",
                kind: Kind::Child {
                    value_kind: "protocol_slot",
                    cardinality: Card::One,
                },
            },
            Field {
                name: "span",
                kind: Kind::Span,
            },
        ],
    },
    Constructor {
        name: "DirectiveBlock",
        value_kind: "statement",
        fields: DIRECTIVE_FIELDS,
    },
    Constructor {
        name: "DirectiveModule",
        value_kind: "statement",
        fields: DIRECTIVE_FIELDS,
    },
];
const DIRECTIVE_FIELDS: &[Field<'static>] = &[
    Field {
        name: "text",
        kind: Kind::Text,
    },
    Field {
        name: "span",
        kind: Kind::Span,
    },
];
const METHOD_FIELDS: &[Field<'static>] = &[
    Field {
        name: "name",
        kind: Kind::Child {
            value_kind: "name",
            cardinality: Card::One,
        },
    },
    Field {
        name: "body",
        kind: Kind::Child {
            value_kind: "body",
            cardinality: Card::One,
        },
    },
    Field {
        name: "span",
        kind: Kind::Span,
    },
];
const FIELD_FIELDS: &[Field<'static>] = &[
    Field {
        name: "name",
        kind: Kind::Child {
            value_kind: "name",
            cardinality: Card::One,
        },
    },
    Field {
        name: "init",
        kind: Kind::Child {
            value_kind: "expression",
            cardinality: Card::Optional,
        },
    },
    Field {
        name: "span",
        kind: Kind::Span,
    },
];
/// Scope policy and its exceptions remain owned AST data during speculation.
const SCOPE_FIELDS: &[Field<'static>] = &[
    Field {
        name: "names",
        kind: Kind::Child {
            value_kind: "name",
            cardinality: Card::Many,
        },
    },
    Field {
        name: "span",
        kind: Kind::Span,
    },
];
pub enum NormalMember {
    Method {
        is_static: bool,
        value: NormalMethod,
    },
    Field {
        is_static: bool,
        value: NormalField,
    },
}
/// Only routing information is retained for undo, never duplicated AST bodies.
pub enum Undo {
    None,
    Class(Vec<MemberRoute>),
}
#[derive(Clone, Copy)]
pub enum MemberRoute {
    Method(bool),
    Field(bool),
}
pub enum Value {
    Expression(Expression),
    Statement(Statement),
    Body(Vec<Statement>),
    Name(String),
    Module(Module),
    NormalMember(NormalMember),
    NormalClass(Box<NormalClass>),
    ProtocolSlot(vybe_ast::ProtocolSlot),
}
pub struct CommonAst;
/// Per-input line index and the complete frontend module policy. Policy is
/// immutable during PEG speculation and survives into the compiler's module.
pub struct Context {
    lines: Vec<usize>,
    pub directives: vybe_ast::Directives,
    /// Explicit frontend spelling registrations, immutable while matching.
    pub protocol_slots: std::collections::HashMap<String, vybe_ast::ProtocolSlot>,
    /// Registered source declarations carry the complete policy delta.
    pub directive_bindings: std::collections::HashMap<String, vybe_ast::Directives>,
}
impl CommonAst {
    pub fn prepare_with_directives(source: &str, directives: vybe_ast::Directives) -> Context {
        Context {
            lines: std::iter::once(0)
                .chain(
                    source
                        .bytes()
                        .enumerate()
                        .filter_map(|(i, c)| (c == b'\n').then_some(i + 1)),
                )
                .collect(),
            directives,
            protocol_slots: Default::default(),
            directive_bindings: Default::default(),
        }
    }
}
fn expression(value: Value) -> Expression {
    let Value::Expression(value) = value else {
        panic!("expression")
    };
    value
}
fn statement(value: Value) -> Statement {
    let Value::Statement(value) = value else {
        panic!("statement")
    };
    value
}
fn body(value: Value) -> Vec<Statement> {
    let Value::Body(value) = value else {
        panic!("body")
    };
    value
}
fn name(value: Value) -> String {
    let Value::Name(value) = value else {
        panic!("name")
    };
    value
}
fn one(arg: Argument<'_, Value>) -> Value {
    let Argument::One(value) = arg else {
        panic!("one")
    };
    value
}
fn many(arg: Argument<'_, Value>) -> Vec<Value> {
    let Argument::Many(value) = arg else {
        panic!("many")
    };
    value
}
fn span(lines: &[usize], bytes: Bytes) -> Span {
    let position = |offset| {
        let line = lines.partition_point(|&start| start <= offset) - 1;
        (line as u32 + 1, (offset - lines[line]) as u32 + 1)
    };
    let (start_line, start_col) = position(bytes.start);
    let (end_line, end_col) = position(bytes.end);
    Span {
        start_line,
        start_col,
        end_line,
        end_col,
    }
}
impl Target for CommonAst {
    type Value = Value;
    type Undo = Undo;
    type Context = Context;
    fn prepare(source: &str) -> Self::Context {
        Self::prepare_with_directives(source, Default::default())
    }
    fn validate(
        context: &Self::Context,
        constructor: usize,
        _: &[grameo::ast::Field],
        span: Bytes,
        text: &str,
    ) -> Result<(), grameo::builder::BuildError> {
        if constructor == C::Integer as usize && text.parse::<i64>().is_err() {
            Err(grameo::builder::BuildError {
                span,
                message: "integer exceeds the common AST range".into(),
            })
        } else if constructor == C::ProtocolSlot as usize
            && !context.protocol_slots.contains_key(text)
            && vybe_ast::ProtocolSlot::from_key(text).is_none()
        {
            Err(grameo::builder::BuildError {
                span,
                message: format!("unregistered protocol slot: {text}"),
            })
        } else if (constructor == C::DirectiveBlock as usize
            || constructor == C::DirectiveModule as usize)
            && !context.directive_bindings.contains_key(text)
        {
            Err(grameo::builder::BuildError {
                span,
                message: format!("unregistered directive: {text}"),
            })
        } else {
            Ok(())
        }
    }
    fn construct(
        lines: &Self::Context,
        ctor: usize,
        bytes: Bytes,
        args: Vec<Argument<'_, Value>>,
    ) -> (Value, Undo) {
        let mut args = args.into_iter();
        let span = span(&lines.lines, bytes);
        let mut undo = Undo::None;
        let value = match ctor {
            x if x == C::Integer as usize => {
                let Argument::Text(text) = args.next().unwrap() else {
                    panic!()
                };
                Value::Expression(Expression::with_span(
                    ExprKind::Lit(Literal::Int(text.parse().unwrap())),
                    span,
                ))
            }
            x if x == C::Name as usize => {
                let Argument::Text(text) = args.next().unwrap() else {
                    panic!()
                };
                Value::Name(text.to_owned())
            }
            x if x == C::DirectiveBlock as usize || x == C::DirectiveModule as usize => {
                let Argument::Text(text) = args.next().unwrap() else {
                    panic!()
                };
                let scope = if x == C::DirectiveBlock as usize {
                    vybe_ast::DirectiveScope::Block
                } else {
                    vybe_ast::DirectiveScope::Module
                };
                Value::Statement(Statement::with_span(
                    StmtKind::Directive {
                        set: lines
                            .directive_bindings
                            .get(text)
                            .expect("validated directive")
                            .clone(),
                        scope,
                    },
                    span,
                ))
            }
            x if x == C::ProtocolSlot as usize => {
                let Argument::Text(text) = args.next().unwrap() else {
                    panic!()
                };
                Value::ProtocolSlot(
                    lines
                        .protocol_slots
                        .get(text)
                        .copied()
                        .or_else(|| vybe_ast::ProtocolSlot::from_key(text))
                        .expect("validated protocol slot"),
                )
            }
            x if x == C::Return as usize => {
                let Argument::Optional(value) = args.next().unwrap() else {
                    panic!()
                };
                Value::Statement(Statement::with_span(
                    StmtKind::Return(value.map(expression)),
                    span,
                ))
            }
            x if x == C::ScopeClosed as usize
                || x == C::ScopeGlobal as usize
                || x == C::ScopeNonlocal as usize =>
            {
                let kind = if x == C::ScopeClosed as usize {
                    ScopeDeclKind::Closed
                } else if x == C::ScopeGlobal as usize {
                    ScopeDeclKind::Global
                } else {
                    ScopeDeclKind::Nonlocal
                };
                let names = many(args.next().unwrap()).into_iter().map(name).collect();
                Value::Statement(Statement::with_span(
                    StmtKind::ScopeDecl { kind, names },
                    span,
                ))
            }
            x if x == C::Block as usize => Value::Body(
                many(args.next().unwrap())
                    .into_iter()
                    .map(statement)
                    .collect(),
            ),
            x if x == C::Statement as usize || x == C::NormalMember as usize => {
                one(args.next().unwrap())
            }
            x if x == C::While as usize => {
                let cond = expression(one(args.next().unwrap()));
                let body = body(one(args.next().unwrap()));
                Value::Statement(Statement::with_span(
                    StmtKind::While {
                        cond,
                        body,
                        else_body: None,
                    },
                    span,
                ))
            }
            x if x == C::Function as usize || x == C::ProtocolFunction as usize => {
                let name = name(one(args.next().unwrap()));
                let body = body(one(args.next().unwrap()));
                let mut modifiers = vybe_ast::Modifiers::default();
                if x == C::ProtocolFunction as usize {
                    let Value::ProtocolSlot(slot) = one(args.next().unwrap()) else {
                        panic!()
                    };
                    modifiers.protocol_slot = Some(slot);
                }
                Value::Statement(Statement::with_span(
                    StmtKind::FunctionDecl {
                        name,
                        body,
                        params: vec![],
                        return_type: None,
                        modifiers,
                        handles: vec![],
                        is_async: false,
                        is_generator: false,
                        is_sub: false,
                    },
                    span,
                ))
            }
            x if x == C::Namespace as usize => {
                let name = name(one(args.next().unwrap()));
                let body = body(one(args.next().unwrap()));
                Value::Statement(Statement::with_span(
                    StmtKind::NamespaceDecl { name, body },
                    span,
                ))
            }
            x if x == C::NormalMethod as usize || x == C::NormalStaticMethod as usize => {
                let name = name(one(args.next().unwrap()));
                let body = body(one(args.next().unwrap()));
                let is_static = x == C::NormalStaticMethod as usize;
                let mut modifiers = vybe_ast::Modifiers::default();
                modifiers.is_static = is_static;
                Value::NormalMember(NormalMember::Method {
                    is_static,
                    value: build_normal_method(
                        span,
                        &name,
                        &name,
                        vec![],
                        None,
                        body,
                        Access::Public,
                        false,
                        false,
                        false,
                        modifiers,
                    ),
                })
            }
            x if x == C::NormalField as usize || x == C::NormalStaticField as usize => {
                let name = name(one(args.next().unwrap()));
                let Argument::Optional(init) = args.next().unwrap() else {
                    panic!()
                };
                Value::NormalMember(NormalMember::Field {
                    is_static: x == C::NormalStaticField as usize,
                    value: NormalField {
                        span,
                        name,
                        type_hint: None,
                        init: init.map(expression),
                        array_bounds: None,
                        access: Access::Public,
                        readonly: false,
                        value_type: None,
                        storage: None,
                    },
                })
            }
            x if x == C::NormalClass as usize => {
                let name = name(one(args.next().unwrap()));
                let members = many(args.next().unwrap());
                let mut routes = Vec::with_capacity(members.len());
                let mut out = NormalMembers::default();
                for member in members {
                    match member {
                        Value::NormalMember(NormalMember::Method { is_static, value }) => {
                            routes.push(MemberRoute::Method(is_static));
                            out.push_method(is_static, value);
                        }
                        Value::NormalMember(NormalMember::Field { is_static, value }) => {
                            routes.push(MemberRoute::Field(is_static));
                            out.push_field(is_static, value);
                        }
                        _ => panic!("normalized class member required"),
                    }
                }
                undo = Undo::Class(routes);
                Value::NormalClass(Box::new(
                    NormalClass {
                        name,
                        span,
                        ..Default::default()
                    }
                    .with_members(out),
                ))
            }
            x if x == C::Class as usize => {
                let name = name(one(args.next().unwrap()));
                let methods = many(args.next().unwrap())
                    .into_iter()
                    .map(|v| ClassMember::Method(Box::new(statement(v))))
                    .collect();
                Value::Statement(Statement::with_span(
                    StmtKind::ClassDecl {
                        name,
                        members: methods,
                        parents: vec![],
                        interfaces: vec![],
                        modifiers: Default::default(),
                        decorators: vec![],
                    },
                    span,
                ))
            }
            x if x == C::Module as usize => Value::Module(Module {
                name: "<module>".into(),
                language: Lang::Unknown,
                body: many(args.next().unwrap())
                    .into_iter()
                    .map(statement)
                    .collect(),
                imports: vec![],
                directives: lines.directives.clone(),
                canon: Default::default(),
            }),
            _ => panic!("unknown generated constructor"),
        };
        (value, undo)
    }
    fn split(ctor: usize, value: Value, undo: Undo) -> Vec<Value> {
        if ctor == C::Statement as usize || ctor == C::NormalMember as usize {
            return vec![value];
        }
        match value {
            Value::Expression(_) | Value::Name(_) | Value::ProtocolSlot(_) => vec![],
            Value::Body(body) => body.into_iter().map(Value::Statement).collect(),
            Value::Module(module) => module.body.into_iter().map(Value::Statement).collect(),
            Value::NormalMember(NormalMember::Method { value, .. }) => {
                vec![Value::Name(value.source_name), Value::Body(value.body)]
            }
            Value::NormalMember(NormalMember::Field { value, .. }) => {
                std::iter::once(Value::Name(value.name))
                    .chain(value.init.map(Value::Expression))
                    .collect()
            }
            Value::NormalClass(class) => {
                let class = *class;
                let Undo::Class(routes) = undo else {
                    panic!("class inverse routing required")
                };
                let mut methods = class.instance_methods.into_iter();
                let mut static_methods = class.static_methods.into_iter();
                let mut fields = class.instance_fields.into_iter();
                let mut static_fields = class.static_fields.into_iter();
                std::iter::once(Value::Name(class.name))
                    .chain(routes.into_iter().map(|route| {
                        Value::NormalMember(match route {
                            MemberRoute::Method(is_static) => NormalMember::Method {
                                is_static,
                                value: if is_static {
                                    static_methods.next()
                                } else {
                                    methods.next()
                                }
                                .unwrap(),
                            },
                            MemberRoute::Field(is_static) => NormalMember::Field {
                                is_static,
                                value: if is_static {
                                    static_fields.next()
                                } else {
                                    fields.next()
                                }
                                .unwrap(),
                            },
                        })
                    }))
                    .collect()
            }
            Value::Statement(stmt) => match stmt.kind {
                StmtKind::Directive { .. } => vec![],
                StmtKind::ScopeDecl { names, .. } => names.into_iter().map(Value::Name).collect(),
                StmtKind::Return(value) => value.into_iter().map(Value::Expression).collect(),
                StmtKind::While { cond, body, .. } => {
                    vec![Value::Expression(cond), Value::Body(body)]
                }
                StmtKind::FunctionDecl {
                    name,
                    body,
                    modifiers,
                    ..
                } => {
                    let mut children = vec![Value::Name(name), Value::Body(body)];
                    if ctor == C::ProtocolFunction as usize {
                        children.push(Value::ProtocolSlot(
                            modifiers.protocol_slot.expect("declared slot"),
                        ));
                    }
                    children
                }
                StmtKind::NamespaceDecl { name, body } => {
                    vec![Value::Name(name), Value::Body(body)]
                }
                StmtKind::ClassDecl { name, members, .. } => std::iter::once(Value::Name(name))
                    .chain(members.into_iter().map(|m| {
                        let ClassMember::Method(method) = m else {
                            panic!()
                        };
                        Value::Statement(*method)
                    }))
                    .collect(),
                _ => panic!("unsupported inverse"),
            },
        }
    }
}

impl Value {
    pub fn into_module(self) -> Module {
        let Self::Module(module) = self else {
            panic!("module root required")
        };
        module
    }
}

#[cfg(test)]
mod scope_tests {
    use super::*;
    use grameo::ast::builder::AstBuilder;
    use grameo::ast::{Binding, Field as BoundField, Source};

    #[test]
    fn grammar_scope_nodes_survive_speculation_in_source_order() {
        let grammar = grameo::compile_with_ast_bindings(
            r#"
// @grameo ast root = Module(body: children*)
root = { statement* ~ EOI }
statement = _{ (closed ~ "!") | closed | global | nonlocal }
// @grameo ast closed = ScopeClosed(names: name*, span: rule)
closed = { "closed" ~ name* ~ ";" }
// @grameo ast global = ScopeGlobal(names: name*, span: rule)
global = { "global" ~ name* ~ ";" }
// @grameo ast nonlocal = ScopeNonlocal(names: name*, span: rule)
nonlocal = { "nonlocal" ~ name* ~ ";" }
// @grameo ast name = Name(text: source)
name = @{ "$"? ~ ASCII_ALPHA+ }
WHITESPACE = _{ " " | "\n" }
"#,
        )
        .unwrap();
        grammar.validate_ast_schema(SCHEMA).unwrap();
        const NAMES: &[BoundField] = &[
            BoundField {
                name: "names",
                source: Source::Child {
                    rule: 5,
                    occurrence: None,
                    cardinality: Card::Many,
                },
            },
            BoundField {
                name: "span",
                source: Source::Span,
            },
        ];
        static BINDINGS: &[Binding] = &[
            Binding {
                rule: 0,
                constructor: C::Module as usize,
                fields: &[BoundField {
                    name: "body",
                    source: Source::Children(Card::Many),
                }],
            },
            Binding {
                rule: 2,
                constructor: C::ScopeClosed as usize,
                fields: NAMES,
            },
            Binding {
                rule: 3,
                constructor: C::ScopeGlobal as usize,
                fields: NAMES,
            },
            Binding {
                rule: 4,
                constructor: C::ScopeNonlocal as usize,
                fields: NAMES,
            },
            Binding {
                rule: 5,
                constructor: C::Name as usize,
                fields: &[BoundField {
                    name: "text",
                    source: Source::Text,
                }],
            },
        ];
        let source = "closed $SERVER $GET;\nglobal x x;\nnonlocal outer;\nclosed;";
        let mut builder = AstBuilder::<CommonAst>::new(BINDINGS, source);
        grameo::engine::build_ast_program(
            &grammar,
            "root",
            source,
            Default::default(),
            &mut builder,
        )
        .unwrap();
        let module = builder.into_value().unwrap().into_module();
        let expected = [
            (ScopeDeclKind::Closed, vec!["$SERVER", "$GET"]),
            (ScopeDeclKind::Global, vec!["x", "x"]),
            (ScopeDeclKind::Nonlocal, vec!["outer"]),
            (ScopeDeclKind::Closed, vec![]),
        ];
        assert_eq!(module.body.len(), expected.len());
        for (i, (stmt, (expected_kind, expected_names))) in
            module.body.iter().zip(expected).enumerate()
        {
            let StmtKind::ScopeDecl { kind, names } = &stmt.kind else {
                panic!("scope node lost")
            };
            assert_eq!(*kind, expected_kind);
            assert_eq!(names, &expected_names);
            assert_eq!(stmt.span.start_line, i as u32 + 1);
        }
    }

    #[test]
    fn configured_module_preserves_complete_policy_through_backtracking() {
        use vybe_ast::{CaseMatch, MethodReceiver, PassBy, ReceiverBinding};
        static BINDINGS: &[Binding] = &[
            Binding {
                rule: 0,
                constructor: C::Module as usize,
                fields: &[BoundField {
                    name: "body",
                    source: Source::Children(Card::Many),
                }],
            },
            Binding {
                rule: 1,
                constructor: C::Return as usize,
                fields: &[
                    BoundField {
                        name: "value",
                        source: Source::Child {
                            rule: 2,
                            occurrence: None,
                            cardinality: Card::Optional,
                        },
                    },
                    BoundField {
                        name: "span",
                        source: Source::Span,
                    },
                ],
            },
            Binding {
                rule: 2,
                constructor: C::Integer as usize,
                fields: &[
                    BoundField {
                        name: "text",
                        source: Source::Text,
                    },
                    BoundField {
                        name: "span",
                        source: Source::Span,
                    },
                ],
            },
        ];
        let grammar = grameo::compile(
            r#"
root = { (statement ~ "!") | statement ~ EOI }
statement = { "return" ~ integer? ~ ";" }
integer = @{ ASCII_DIGIT+ }
WHITESPACE = _{ " " }
"#,
        )
        .unwrap();
        let policy = vybe_ast::Directives {
            reference_binding: Some(PassBy::Alias),
            receiver_binding: Some(ReceiverBinding::UniversalParameter),
            method_receiver: Some(MethodReceiver::CallSite),
            variable_case: Some(CaseMatch::Exact),
            callable_case: Some(CaseMatch::Folded),
            instance_fields_are_own_properties: Some(false),
            ..Default::default()
        };
        let source = "return 7;";
        let context = CommonAst::prepare_with_directives(source, policy.clone());
        let mut builder = AstBuilder::<CommonAst>::with_context(BINDINGS, context);
        grameo::engine::build_ast_program(
            &grammar,
            "root",
            source,
            Default::default(),
            &mut builder,
        )
        .unwrap();
        let module = builder.into_value().unwrap().into_module();
        assert_eq!(module.directives, policy);
        assert_eq!(module.body.len(), 1);
        assert!(matches!(module.body[0].kind, StmtKind::Return(Some(_))));
    }

    #[test]
    fn scope_inverse_returns_original_name_allocations() {
        for ctor in [C::ScopeClosed, C::ScopeGlobal, C::ScopeNonlocal] {
            let original = String::from("outer");
            let pointer = original.as_ptr();
            let lines = CommonAst::prepare("outer");
            let (value, undo) = CommonAst::construct(
                &lines,
                ctor as usize,
                Bytes { start: 0, end: 5 },
                vec![
                    Argument::Many(vec![Value::Name(original)]),
                    Argument::Span(Bytes { start: 0, end: 5 }),
                ],
            );
            let mut children = CommonAst::split(ctor as usize, value, undo);
            let Value::Name(restored) = children.pop().unwrap() else {
                panic!()
            };
            assert_eq!(restored.as_ptr(), pointer);
        }
    }
}

#[cfg(test)]
mod protocol_hook_tests {
    use super::*;
    use grameo::ast::builder::AstBuilder;
    use grameo::ast::{Binding, Field as BoundField, Source};
    static BINDINGS: &[Binding] = &[
        Binding {
            rule: 0,
            constructor: C::Module as usize,
            fields: &[BoundField {
                name: "body",
                source: Source::Children(Card::Many),
            }],
        },
        Binding {
            rule: 1,
            constructor: C::Class as usize,
            fields: &[
                BoundField {
                    name: "name",
                    source: Source::Child {
                        rule: 2,
                        occurrence: None,
                        cardinality: Card::One,
                    },
                },
                BoundField {
                    name: "methods",
                    source: Source::Child {
                        rule: 3,
                        occurrence: None,
                        cardinality: Card::Many,
                    },
                },
                BoundField {
                    name: "span",
                    source: Source::Span,
                },
            ],
        },
        Binding {
            rule: 2,
            constructor: C::Name as usize,
            fields: &[BoundField {
                name: "text",
                source: Source::Text,
            }],
        },
        Binding {
            rule: 3,
            constructor: C::ProtocolFunction as usize,
            fields: &[
                BoundField {
                    name: "name",
                    source: Source::Child {
                        rule: 2,
                        occurrence: None,
                        cardinality: Card::One,
                    },
                },
                BoundField {
                    name: "body",
                    source: Source::Child {
                        rule: 5,
                        occurrence: None,
                        cardinality: Card::One,
                    },
                },
                BoundField {
                    name: "slot",
                    source: Source::Child {
                        rule: 4,
                        occurrence: None,
                        cardinality: Card::One,
                    },
                },
                BoundField {
                    name: "span",
                    source: Source::Span,
                },
            ],
        },
        Binding {
            rule: 4,
            constructor: C::ProtocolSlot as usize,
            fields: &[BoundField {
                name: "text",
                source: Source::Text,
            }],
        },
        Binding {
            rule: 5,
            constructor: C::Block as usize,
            fields: &[BoundField {
                name: "body",
                source: Source::Children(Card::Many),
            }],
        },
        Binding {
            rule: 6,
            constructor: C::Return as usize,
            fields: &[
                BoundField {
                    name: "value",
                    source: Source::Child {
                        rule: 7,
                        occurrence: None,
                        cardinality: Card::Optional,
                    },
                },
                BoundField {
                    name: "span",
                    source: Source::Span,
                },
            ],
        },
        Binding {
            rule: 7,
            constructor: C::Integer as usize,
            fields: &[
                BoundField {
                    name: "text",
                    source: Source::Text,
                },
                BoundField {
                    name: "span",
                    source: Source::Span,
                },
            ],
        },
    ];
    const GRAMMAR: &str = r#"
root = { (class ~ "!") | class ~ EOI }
class = { "class" ~ name ~ "{" ~ method* ~ "}" }
name = @{ (ASCII_ALPHA | "_")+ }
method = { slot ~ "fn" ~ name ~ body }
slot = @{ (ASCII_ALPHA | "_")+ }
body = { "{" ~ statement* ~ "}" }
statement = { "return" ~ integer? ~ ";" }
integer = @{ ASCII_DIGIT+ }
WHITESPACE = _{ " " | NEWLINE }
"#;
    #[test]
    fn registered_slot_hooks_preserve_method_identity_and_rollback() {
        let grammar = grameo::compile(GRAMMAR).unwrap();
        let source = "class C {\n __toString fn render { return 7; }\n}";
        let mut context = CommonAst::prepare_with_directives(source, Default::default());
        context
            .protocol_slots
            .insert("__toString".into(), vybe_ast::ProtocolSlot::ToString);
        let mut builder = AstBuilder::<CommonAst>::with_context(BINDINGS, context);
        grameo::engine::build_ast_program(
            &grammar,
            "root",
            source,
            Default::default(),
            &mut builder,
        )
        .unwrap();
        let module = builder.into_value().unwrap().into_module();
        let StmtKind::ClassDecl { members, .. } = &module.body[0].kind else {
            panic!()
        };
        assert_eq!(
            vybe_ast::class_normalize::declared_protocol_slots(members),
            vec![vybe_ast::ProtocolSlot::ToString]
        );
        let ClassMember::Method(method) = &members[0] else {
            panic!()
        };
        let StmtKind::FunctionDecl {
            name,
            body,
            modifiers,
            ..
        } = &method.kind
        else {
            panic!()
        };
        assert_eq!(name, "render");
        assert_eq!(
            modifiers.protocol_slot,
            Some(vybe_ast::ProtocolSlot::ToString)
        );
        assert_eq!(body.len(), 1);
        assert_eq!(method.span.start_line, 2);
    }
    #[test]
    fn all_canonical_slots_construct_and_restore_the_existing_ast_role() {
        let grammar = grameo::compile(GRAMMAR).unwrap();
        for slot in vybe_ast::ProtocolSlot::ALL {
            let source = format!("class C {{ {} fn render {{ return 7; }} }}", slot.as_key());
            let mut builder = AstBuilder::<CommonAst>::new(BINDINGS, &source);
            grameo::engine::build_ast_program(
                &grammar,
                "root",
                &source,
                Default::default(),
                &mut builder,
            )
            .unwrap();
            let module = builder.into_value().unwrap().into_module();
            let StmtKind::ClassDecl { members, .. } = &module.body[0].kind else {
                panic!()
            };
            assert_eq!(
                vybe_ast::class_normalize::declared_protocol_slots(members),
                vec![slot]
            );
        }
    }

    #[test]
    fn absent_hook_keeps_peg_valid_but_explicit_unknown_slot_fails() {
        let grammar = grameo::compile(GRAMMAR).unwrap();
        let source = "class C { unknown fn render { return 7; } }";
        assert!(grammar.parse_source("root", source).is_ok());
        let mut builder = AstBuilder::<CommonAst>::new(BINDINGS, source);
        assert!(
            grameo::engine::build_ast_program(
                &grammar,
                "root",
                source,
                Default::default(),
                &mut builder
            )
            .is_err()
        );
    }
}

#[cfg(test)]
mod directive_hook_tests {
    use super::*;
    use grameo::ast::builder::AstBuilder;
    use grameo::ast::{Binding, Field as BoundField, Source};
    #[test]
    fn positional_directive_hooks_preserve_full_deltas_and_scope_on_rollback() {
        static FIELDS: &[BoundField] = &[
            BoundField {
                name: "text",
                source: Source::Text,
            },
            BoundField {
                name: "span",
                source: Source::Span,
            },
        ];
        static BINDINGS: &[Binding] = &[
            Binding {
                rule: 0,
                constructor: C::Module as usize,
                fields: &[BoundField {
                    name: "body",
                    source: Source::Children(Card::Many),
                }],
            },
            Binding {
                rule: 1,
                constructor: C::DirectiveBlock as usize,
                fields: FIELDS,
            },
            Binding {
                rule: 2,
                constructor: C::DirectiveModule as usize,
                fields: FIELDS,
            },
        ];
        let grammar = grameo::compile(
            r#"root = { (block ~ module ~ "!") | block ~ module ~ EOI }
block = @{ "local" }
module = @{ "module" }
WHITESPACE = _{ " " }
"#,
        )
        .unwrap();
        let source = "local module";
        let policy = vybe_ast::Directives {
            reference_binding: Some(vybe_ast::PassBy::Alias),
            receiver_binding: Some(vybe_ast::ReceiverBinding::UniversalParameter),
            method_receiver: Some(vybe_ast::MethodReceiver::CallSite),
            instance_fields_are_own_properties: Some(false),
            ..Default::default()
        };
        let mut context = CommonAst::prepare_with_directives(source, Default::default());
        context
            .directive_bindings
            .insert("local".into(), policy.clone());
        context
            .directive_bindings
            .insert("module".into(), policy.clone());
        let mut builder = AstBuilder::<CommonAst>::with_context(BINDINGS, context);
        grameo::engine::build_ast_program(
            &grammar,
            "root",
            source,
            Default::default(),
            &mut builder,
        )
        .unwrap();
        let module = builder.into_value().unwrap().into_module();
        assert_eq!(module.directives, vybe_ast::Directives::default());
        for (stmt, expected_scope) in module.body.iter().zip([
            vybe_ast::DirectiveScope::Block,
            vybe_ast::DirectiveScope::Module,
        ]) {
            let StmtKind::Directive { set, scope } = &stmt.kind else {
                panic!()
            };
            assert_eq!(*scope, expected_scope);
            assert_eq!(*set, policy);
        }
        assert_eq!(module.body.len(), 2);
        assert_eq!(module.body[1].span.start_col, 7);
        assert!(grammar.parse_source("root", source).is_ok());
        let mut absent = AstBuilder::<CommonAst>::new(BINDINGS, source);
        assert!(
            grameo::engine::build_ast_program(
                &grammar,
                "root",
                source,
                Default::default(),
                &mut absent
            )
            .is_err()
        );
    }
}
