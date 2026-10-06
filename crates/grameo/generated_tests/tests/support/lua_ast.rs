//! Direct binding for the unchanged Lua grammar, with ordinary and native-Pratt
//! execution preserving the same AST/span policies. Final dialect normalization
//! remains required before production promotion.
#[path = "lua_numeric.rs"]
mod numeric;
use vybe_ast::*;
use grameo::{
    CaptureRule, Span,
    arena::{Id, MetadataArena},
    builder::{BuildError, Builder},
};
use vybe_parser_generated_tests::lua::Rule;

#[allow(clippy::manual_strip, clippy::collapsible_if)]
mod literals {
    include!(concat!(env!("OUT_DIR"), "/lua_literal_helpers.rs"));
}
enum Value {
    Expression(Expression),
    Operator(Operator),
    Field(Box<ArrayElement>),
    Param(Box<Param>),
    Params(Vec<Param>),
    Statement(Box<Statement>),
    Body(Vec<Statement>),
    Expressions(Vec<Expression>),
    Module(Box<Module>),
}
#[derive(Clone, Copy)]
enum Operator {
    Binary(BinOp),
    Unary(UnaryOp),
    Length,
}
enum Metadata {
    Span(vybe_ast::Span),
    Ident { name: String, span: vybe_ast::Span },
    ForNames(Vec<Expression>),
    InsertedSelf(bool),
    NumericPromotion(Vec<(usize, i64)>),
    NumericCapture(numeric::CaptureJournal),
}
#[derive(Clone, Copy)]
struct Frame {
    rule: Rule,
    base: usize,
    field: Option<FieldKind>,
    names: usize,
    method: bool,
    context: usize,
}
#[derive(Clone, Copy)]
enum FieldKind {
    Named,
    Computed,
    Spread,
}
enum Undo {
    Arena(usize),
    PopFrame,
    RestoreFrame(Frame),
    RestoreFrameState { index: usize, frame: Frame },
    RemovePrattOrigin(Id),
}
pub struct Ast<'a> {
    source: &'a str,
    pratt: bool,
    pratt_origins: std::collections::HashMap<Id, (usize, u8)>,
    // Existing common-AST spans use LF-only lines and byte columns. Preserve
    // that policy here; editor coordinates use source::Index separately.
    line_starts: Vec<usize>,
    values: MetadataArena<Value, Metadata>,
    frames: Vec<Frame>,
    undo: Vec<Undo>,
}
fn expression(value: Value) -> Expression {
    let Value::Expression(value) = value else {
        unreachable!("expected Lua expression")
    };
    value
}
fn rule(id: usize) -> Rule {
    Rule::from_capture(CaptureRule::User(id)).unwrap()
}
fn wrapper(rule: Rule) -> bool {
    matches!(
        rule,
        Rule::expr
            | Rule::chunk
            | Rule::or_expr
            | Rule::and_expr
            | Rule::compare_expr
            | Rule::bor_expr
            | Rule::bxor_expr
            | Rule::band_expr
            | Rule::shift_expr
            | Rule::concat_expr
            | Rule::additive
            | Rule::multiplicative
            | Rule::unary
            | Rule::pow_expr
            | Rule::postfix
            | Rule::call_expression
            | Rule::primary
            | Rule::long_string
            | Rule::call_chain
            | Rule::table_field
            | Rule::table_constructor
            | Rule::param
            | Rule::param_list
            | Rule::func_expr
            | Rule::func_body
            | Rule::body_block
            | Rule::then_block
            | Rule::else_block
            | Rule::block
            | Rule::return_statement
            | Rule::local_var
            | Rule::local_function
            | Rule::function_decl
            | Rule::func_name
            | Rule::func_name_suffix
            | Rule::for_numeric
            | Rule::for_generic
            | Rule::for_varlist
            | Rule::do_statement
            | Rule::if_statement
            | Rule::while_statement
            | Rule::repeat_statement
            | Rule::assign_stmt
            | Rule::assign_lhs
            | Rule::assign_rhs
            | Rule::label_statement
            | Rule::goto_statement
    )
}
impl<'a> Ast<'a> {
    pub fn new(source: &'a str) -> Self {
        Self {
            source,
            pratt: false,
            pratt_origins: Default::default(),
            line_starts: std::iter::once(0)
                .chain(source.match_indices('\n').map(|(offset, _)| offset + 1))
                .collect(),
            values: Default::default(),
            frames: Vec::new(),
            undo: Vec::new(),
        }
    }
    pub fn with_pratt(source: &'a str) -> Self {
        Self {
            pratt: true,
            ..Self::new(source)
        }
    }
    pub fn into_expression(self) -> Expression {
        assert!(self.frames.is_empty());
        assert_eq!(self.values.values().len(), 1);
        let root = self.values.values()[0];
        expression(self.values.into_root(root))
    }
    pub fn into_module(self) -> Module {
        assert!(self.frames.is_empty());
        assert_eq!(self.values.values().len(), 1);
        let root = self.values.values()[0];
        let Value::Module(module) = self.values.into_root(root) else {
            unreachable!("expected Lua module")
        };
        *module
    }
    pub fn is_empty(&self) -> bool {
        self.frames.is_empty()
            && self.values.values().is_empty()
            && self.undo.is_empty()
            && self.pratt_origins.is_empty()
    }
    fn span(&self, span: Span) -> vybe_ast::Span {
        let at = |offset| {
            let line = self.line_starts.partition_point(|start| *start <= offset) - 1;
            (
                line as u32 + 1,
                (offset - self.line_starts[line]) as u32 + 1,
            )
        };
        let a = at(span.start);
        let b = at(span.end);
        vybe_ast::Span {
            start_line: a.0,
            start_col: a.1,
            end_line: b.0,
            end_col: b.1,
        }
    }
    fn push(&mut self, value: Value) {
        let id = self.values.alloc(value);
        self.values.push(id);
    }
    fn set_span(&mut self, id: Id, span: vybe_ast::Span) {
        let Value::Expression(value) = self.values.get(id) else {
            unreachable!()
        };
        let old = value.span;
        if (old.start_line, old.start_col, old.end_line, old.end_col)
            == (span.start_line, span.start_col, span.end_line, span.end_col)
        {
            return;
        }
        self.values.edit(
            id,
            |value| {
                let Value::Expression(value) = value else {
                    unreachable!()
                };
                Metadata::Span(std::mem::replace(&mut value.span, span))
            },
            |value, previous| {
                let (Value::Expression(value), Metadata::Span(previous)) = (value, previous) else {
                    unreachable!()
                };
                value.span = previous;
            },
        );
    }
    fn flat_chain_child(&mut self, id: Id, op: BinOp) {
        let context = self.frames.last().unwrap().context;
        if self
            .pratt_origins
            .get(&id)
            .is_some_and(|&(origin, tier)| origin == context && Some(tier) == lua_tier(op))
        {
            // Only reductions in this expression's flat tier lose their span.
            // Parenthesized/nested rule roots carry a different context.
            self.set_span(id, Default::default());
        }
    }
    fn binary(&mut self, left: Id, right: Id, op: BinOp) -> Id {
        self.values.binary(
            left,
            right,
            |left, right| {
                Value::Expression(Expression::new(ExprKind::Binary {
                    op,
                    left: Box::new(expression(left)),
                    right: Box::new(expression(right)),
                }))
            },
            |value| {
                let ExprKind::Binary { left, right, .. } = expression(value).kind else {
                    unreachable!()
                };
                (Value::Expression(*left), Value::Expression(*right))
            },
        )
    }
    fn collect(&mut self, base: usize) -> Vec<Id> {
        let mut children = Vec::new();
        while self.values.values().len() > base {
            children.push(self.values.pop().unwrap());
        }
        children.reverse();
        children
    }
    fn table_field(&mut self, frame: Frame) {
        let children = self.collect(frame.base);
        let field = if matches!(frame.field, Some(FieldKind::Named | FieldKind::Computed)) {
            assert_eq!(children.len(), 2);
            let mut key = children[0];
            if matches!(frame.field, Some(FieldKind::Named)) {
                self.set_span(key, Default::default());
                key = self.values.unary(
                    key,
                    |value| {
                        let ExprKind::Ident(name) = expression(value).kind else {
                            unreachable!()
                        };
                        Value::Expression(Expression::new(ExprKind::Lit(Literal::Str(name))))
                    },
                    |value| {
                        let ExprKind::Lit(Literal::Str(name)) = expression(value).kind else {
                            unreachable!()
                        };
                        Value::Expression(Expression::new(ExprKind::Ident(name)))
                    },
                );
            }
            self.values.binary(
                key,
                children[1],
                |key, value| {
                    Value::Field(Box::new(ArrayElement {
                        key: Some(expression(key)),
                        value: expression(value),
                        spread: false,
                        by_ref: false,
                    }))
                },
                |value| {
                    let Value::Field(field) = value else {
                        unreachable!()
                    };
                    (
                        Value::Expression(field.key.unwrap()),
                        Value::Expression(field.value),
                    )
                },
            )
        } else {
            assert_eq!(children.len(), 1);
            let spread = matches!(frame.field, Some(FieldKind::Spread));
            self.values.unary(
                children[0],
                |value| {
                    Value::Field(Box::new(ArrayElement {
                        key: None,
                        value: expression(value),
                        spread,
                        by_ref: false,
                    }))
                },
                |value| {
                    let Value::Field(field) = value else {
                        unreachable!()
                    };
                    Value::Expression(field.value)
                },
            )
        };
        self.values.push(field);
    }
    fn table_constructor(&mut self, base: usize) {
        let children = self.collect(base);
        let table = self.values.list(
            children,
            |values| {
                let fields = values
                    .into_iter()
                    .map(|value| {
                        let Value::Field(field) = value else {
                            unreachable!()
                        };
                        *field
                    })
                    .collect();
                Value::Expression(Expression::new(ExprKind::Array(fields)))
            },
            |value| {
                let ExprKind::Array(fields) = expression(value).kind else {
                    unreachable!()
                };
                fields
                    .into_iter()
                    .map(|field| Value::Field(Box::new(field)))
                    .collect()
            },
        );
        self.values.push(table);
    }
    fn mark_field(&mut self, kind: FieldKind) {
        if let Some(frame) = self.frames.last().copied()
            && frame.rule == Rule::table_field
        {
            let index = self.frames.len() - 1;
            self.undo.push(Undo::RestoreFrameState { index, frame });
            self.frames[index].field = Some(kind);
        }
    }
    fn mark_local_name(&mut self) {
        if let Some(frame) = self.frames.last().copied()
            && frame.rule == Rule::local_var
        {
            let index = self.frames.len() - 1;
            self.undo.push(Undo::RestoreFrameState { index, frame });
            self.frames[index].names += 1;
        }
    }
    fn parameter(&mut self, frame: Frame, rest: bool) {
        if rest {
            self.push(Value::Expression(Expression::new(ExprKind::Ident(
                "_lua_varargs".into(),
            ))));
        }
        let children = self.collect(frame.base);
        assert_eq!(children.len(), 1);
        self.set_span(children[0], Default::default());
        let parameter = self.values.unary(
            children[0],
            |value| {
                let ExprKind::Ident(name) = expression(value).kind else {
                    unreachable!()
                };
                Value::Param(Box::new(Param {
                    name,
                    type_hint: None,
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: rest,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                }))
            },
            |value| {
                let Value::Param(parameter) = value else {
                    unreachable!()
                };
                Value::Expression(Expression::new(ExprKind::Ident(parameter.name)))
            },
        );
        self.values.push(parameter);
    }
    fn parameters(&mut self, base: usize) {
        let children = self.collect(base);
        let parameters = self.values.list(
            children,
            |values| {
                Value::Params(
                    values
                        .into_iter()
                        .map(|value| {
                            let Value::Param(value) = value else {
                                unreachable!()
                            };
                            *value
                        })
                        .collect(),
                )
            },
            |value| {
                let Value::Params(parameters) = value else {
                    unreachable!()
                };
                parameters
                    .into_iter()
                    .map(|parameter| Value::Param(Box::new(parameter)))
                    .collect()
            },
        );
        self.values.push(parameters);
    }
    fn body(&mut self, base: usize) {
        let children = self.collect(base);
        let body = self.values.list(
            children,
            |values| {
                Value::Body(
                    values
                        .into_iter()
                        .map(|value| {
                            let Value::Statement(value) = value else {
                                unreachable!("unsupported Lua statement reached body")
                            };
                            *value
                        })
                        .collect(),
                )
            },
            |value| {
                let Value::Body(body) = value else {
                    unreachable!()
                };
                body.into_iter()
                    .map(|statement| Value::Statement(Box::new(statement)))
                    .collect()
            },
        );
        self.values.push(body);
    }
    fn expression_list(&mut self, base: usize) {
        let children = self.collect(base);
        let list = self.values.list(
            children,
            |values| Value::Expressions(values.into_iter().map(expression).collect()),
            |value| {
                let Value::Expressions(values) = value else {
                    unreachable!()
                };
                values.into_iter().map(Value::Expression).collect()
            },
        );
        self.values.push(list);
    }
    fn for_names(&mut self, base: usize) {
        self.expression_list(base);
        let names = *self.values.values().last().unwrap();
        // The legacy ForIn ABI encodes extra names in one string. Keep the
        // exact owned names in typed undo rather than trying to reverse that
        // encoding (a legal identifier can itself contain its separator).
        self.values.edit(
            names,
            |value| {
                let Value::Expressions(names) = value else {
                    unreachable!()
                };
                let previous = std::mem::take(names);
                fn spelling(value: &Expression) -> &str {
                    let ExprKind::Ident(name) = &value.kind else {
                        unreachable!()
                    };
                    name.as_str()
                }
                names.push(Expression::new(ExprKind::Ident(
                    spelling(&previous[0]).to_owned(),
                )));
                if previous.len() > 1 {
                    names.push(Expression::new(ExprKind::Ident(
                        previous[1..]
                            .iter()
                            .map(spelling)
                            .collect::<Vec<_>>()
                            .join("__lua_extra__"),
                    )));
                }
                Metadata::ForNames(previous)
            },
            |value, previous| {
                let (Value::Expressions(names), Metadata::ForNames(previous)) = (value, previous)
                else {
                    unreachable!()
                };
                *names = previous;
            },
        );
    }
    fn numeric_for(&mut self, frame: Frame, span: vybe_ast::Span) {
        let mut children = self.collect(frame.base);
        assert!(matches!(children.len(), 4 | 5));
        if children.len() == 4 {
            let step = self.values.alloc(Value::Expression(Expression::int(1)));
            children.insert(3, step);
        }
        let name = children[0];
        self.set_span(name, Default::default());
        let floating = children[1..4].iter().any(|&id| {
            let Value::Expression(expr) = self.values.get(id) else {
                unreachable!()
            };
            numeric::contains_float(expr)
        });
        if floating {
            for &id in &children[1..4] {
                self.values.edit(
                    id,
                    |value| {
                        let Value::Expression(expr) = value else {
                            unreachable!()
                        };
                        Metadata::NumericPromotion(numeric::promote(expr))
                    },
                    |value, previous| {
                        let (Value::Expression(expr), Metadata::NumericPromotion(journal)) =
                            (value, previous)
                        else {
                            unreachable!()
                        };
                        numeric::restore_promotion(expr, journal);
                    },
                );
            }
        }
        let Value::Expression(Expression {
            kind: ExprKind::Ident(var),
            ..
        }) = self.values.get(name)
        else {
            unreachable!()
        };
        let var = var.to_owned();
        self.values.edit(
            children[4],
            |value| {
                let Value::Body(body) = value else {
                    unreachable!()
                };
                Metadata::NumericCapture(numeric::capture(body, &var))
            },
            |value, previous| {
                let (Value::Body(body), Metadata::NumericCapture(journal)) = (value, previous)
                else {
                    unreachable!()
                };
                numeric::restore_capture(body, journal);
            },
        );
        let statement = self.values.list(
            children,
            |values| {
                let mut values = values.into_iter();
                let ExprKind::Ident(var) = expression(values.next().unwrap()).kind else {
                    unreachable!()
                };
                let start = expression(values.next().unwrap());
                let limit = expression(values.next().unwrap());
                let step = expression(values.next().unwrap());
                let Value::Body(body) = values.next().unwrap() else {
                    unreachable!()
                };
                Value::Statement(Box::new(Statement::with_span(
                    numeric::build(&var, start, limit, step, body),
                    span,
                )))
            },
            |value| {
                let Value::Statement(statement) = value else {
                    unreachable!()
                };
                let StmtKind::For {
                    init: Some(init),
                    mut body,
                    ..
                } = statement.kind
                else {
                    unreachable!()
                };
                let StmtKind::VarDecl { declarations, .. } = init.kind else {
                    unreachable!()
                };
                let StmtKind::Block(mut body) = body.remove(0).kind else {
                    unreachable!()
                };
                let StmtKind::VarDecl {
                    declarations: mut binding,
                    ..
                } = body.remove(0).kind
                else {
                    unreachable!()
                };
                let BindingPattern::Ident(var) = binding.remove(0).pattern else {
                    unreachable!()
                };
                let mut values = vec![Value::Expression(Expression::new(ExprKind::Ident(var)))];
                values.extend(
                    declarations
                        .into_iter()
                        .map(|decl| Value::Expression(decl.init.unwrap())),
                );
                values.push(Value::Body(body));
                values
            },
        );
        self.values.push(statement);
    }
    fn generic_for(&mut self, frame: Frame, span: vybe_ast::Span) {
        let mut children = self.collect(frame.base);
        let names = children.remove(0);
        let body = children.pop().unwrap();
        assert!(!children.is_empty());
        let iter = if children.len() == 1 {
            children[0]
        } else {
            self.values.list(
                children,
                |values| {
                    Value::Expression(Expression::new(ExprKind::Sequence(
                        values.into_iter().map(expression).collect(),
                    )))
                },
                |value| {
                    let ExprKind::Sequence(values) = expression(value).kind else {
                        unreachable!()
                    };
                    values.into_iter().map(Value::Expression).collect()
                },
            )
        };
        let statement = self.values.list(
            vec![names, iter, body],
            |values| {
                let mut values = values.into_iter();
                let Value::Expressions(names) = values.next().unwrap() else {
                    unreachable!()
                };
                let mut names = names.into_iter().map(|name| {
                    let ExprKind::Ident(name) = name.kind else {
                        unreachable!()
                    };
                    name
                });
                let var = names.next().unwrap();
                let key = names.next();
                assert!(names.next().is_none());
                let iter = expression(values.next().unwrap());
                let Value::Body(body) = values.next().unwrap() else {
                    unreachable!()
                };
                let body = if body.is_empty() {
                    body
                } else {
                    vec![Statement::new(StmtKind::Block(body))]
                };
                Value::Statement(Box::new(Statement::with_span(
                    StmtKind::ForIn {
                        var,
                        key,
                        iter,
                        body,
                        of: true,
                        else_body: None,
                        is_async: false,
                    },
                    span,
                )))
            },
            |value| {
                let Value::Statement(statement) = value else {
                    unreachable!()
                };
                let StmtKind::ForIn {
                    var,
                    key,
                    iter,
                    mut body,
                    ..
                } = statement.kind
                else {
                    unreachable!()
                };
                let mut names = vec![Expression::new(ExprKind::Ident(var))];
                if let Some(key) = key {
                    names.push(Expression::new(ExprKind::Ident(key)));
                }
                let body = if body.is_empty() {
                    body
                } else {
                    assert_eq!(body.len(), 1);
                    let StmtKind::Block(body) = body.pop().unwrap().kind else {
                        unreachable!()
                    };
                    body
                };
                vec![
                    Value::Expressions(names),
                    Value::Expression(iter),
                    Value::Body(body),
                ]
            },
        );
        self.values.push(statement);
    }
    fn assignment(&mut self, frame: Frame, span: vybe_ast::Span) {
        let children = self.collect(frame.base);
        assert_eq!(children.len(), 2);
        let Value::Expressions(rhs) = self.values.get(children[1]) else {
            unreachable!()
        };
        let single = rhs.len() == 1;
        let value = self.values.unary(
            children[1],
            |value| {
                let Value::Expressions(mut values) = value else {
                    unreachable!()
                };
                let expression = if single {
                    values.pop().unwrap()
                } else {
                    Expression::new(ExprKind::Array(
                        values
                            .into_iter()
                            .map(|value| ArrayElement {
                                key: None,
                                value,
                                spread: false,
                                by_ref: false,
                            })
                            .collect(),
                    ))
                };
                Value::Expression(expression)
            },
            if single {
                split_single_rhs
            } else {
                split_multiple_rhs
            },
        );
        let statement = self.values.binary(
            children[0],
            value,
            |targets, value| {
                let Value::Expressions(targets) = targets else {
                    unreachable!()
                };
                Value::Statement(Box::new(Statement::with_span(
                    StmtKind::Assign {
                        targets,
                        value: expression(value),
                        by_ref: false,
                    },
                    span,
                )))
            },
            |value| {
                let Value::Statement(statement) = value else {
                    unreachable!()
                };
                let StmtKind::Assign { targets, value, .. } = statement.kind else {
                    unreachable!()
                };
                (Value::Expressions(targets), Value::Expression(value))
            },
        );
        self.values.push(statement);
    }
    fn control_statement(&mut self, frame: Frame, span: vybe_ast::Span) {
        let children = self.collect(frame.base);
        let statement = match frame.rule {
            Rule::do_statement => {
                assert_eq!(children.len(), 1);
                self.values.unary(
                    children[0],
                    |value| {
                        let Value::Body(body) = value else {
                            unreachable!()
                        };
                        Value::Statement(Box::new(Statement::with_span(
                            StmtKind::Block(body),
                            span,
                        )))
                    },
                    |value| {
                        let Value::Statement(statement) = value else {
                            unreachable!()
                        };
                        let StmtKind::Block(body) = statement.kind else {
                            unreachable!()
                        };
                        Value::Body(body)
                    },
                )
            }
            Rule::while_statement => {
                assert_eq!(children.len(), 2);
                self.values.binary(
                    children[0],
                    children[1],
                    |cond, body| {
                        let Value::Body(body) = body else {
                            unreachable!()
                        };
                        let body = if body.is_empty() {
                            body
                        } else {
                            vec![Statement::new(StmtKind::Block(body))]
                        };
                        Value::Statement(Box::new(Statement::with_span(
                            StmtKind::While {
                                cond: expression(cond),
                                body,
                                else_body: None,
                            },
                            span,
                        )))
                    },
                    |value| {
                        let Value::Statement(statement) = value else {
                            unreachable!()
                        };
                        let StmtKind::While { cond, mut body, .. } = statement.kind else {
                            unreachable!()
                        };
                        let body = if body.is_empty() {
                            body
                        } else {
                            assert_eq!(body.len(), 1);
                            let StmtKind::Block(body) = body.pop().unwrap().kind else {
                                unreachable!()
                            };
                            body
                        };
                        (Value::Expression(cond), Value::Body(body))
                    },
                )
            }
            Rule::repeat_statement => {
                assert_eq!(children.len(), 2);
                self.values.binary(
                    children[0],
                    children[1],
                    |body, cond| {
                        let Value::Body(body) = body else {
                            unreachable!()
                        };
                        Value::Statement(Box::new(Statement::with_span(
                            StmtKind::DoWhile {
                                body,
                                cond: expression(cond),
                                until: true,
                            },
                            span,
                        )))
                    },
                    |value| {
                        let Value::Statement(statement) = value else {
                            unreachable!()
                        };
                        let StmtKind::DoWhile { body, cond, .. } = statement.kind else {
                            unreachable!()
                        };
                        (Value::Body(body), Value::Expression(cond))
                    },
                )
            }
            _ => unreachable!(),
        };
        self.values.push(statement);
    }
    fn conditional(&mut self, frame: Frame, span: vybe_ast::Span) {
        let children = self.collect(frame.base);
        let statement = self.values.list(
            children,
            |values| {
                let mut values = values.into_iter();
                let cond = expression(values.next().unwrap());
                let Value::Body(then_body) = values.next().unwrap() else {
                    unreachable!()
                };
                let mut elifs = Vec::new();
                let mut else_body = None;
                while let Some(value) = values.next() {
                    match value {
                        Value::Expression(cond) => {
                            let Value::Body(body) = values.next().unwrap() else {
                                unreachable!()
                            };
                            elifs.push((cond, body));
                        }
                        Value::Body(body) => {
                            else_body = Some(body);
                            assert!(values.next().is_none());
                        }
                        _ => unreachable!(),
                    }
                }
                Value::Statement(Box::new(Statement::with_span(
                    StmtKind::If {
                        cond,
                        then_body,
                        elifs,
                        else_body,
                    },
                    span,
                )))
            },
            |value| {
                let Value::Statement(statement) = value else {
                    unreachable!()
                };
                let StmtKind::If {
                    cond,
                    then_body,
                    elifs,
                    else_body,
                } = statement.kind
                else {
                    unreachable!()
                };
                let mut values = vec![Value::Expression(cond), Value::Body(then_body)];
                for (cond, body) in elifs {
                    values.push(Value::Expression(cond));
                    values.push(Value::Body(body));
                }
                if let Some(body) = else_body {
                    values.push(Value::Body(body));
                }
                values
            },
        );
        self.values.push(statement);
    }
    fn expression_statement(&mut self, rule: Rule) {
        if rule != Rule::expr
            || !self.frames.last().is_some_and(|frame| {
                matches!(
                    frame.rule,
                    Rule::func_body
                        | Rule::body_block
                        | Rule::then_block
                        | Rule::else_block
                        | Rule::block
                )
            })
        {
            return;
        }
        let operand = self.values.pop().unwrap();
        let statement = self.values.unary(
            operand,
            |value| {
                let value = expression(value);
                let span = value.span;
                Value::Statement(Box::new(Statement::with_span(StmtKind::Expr(value), span)))
            },
            |value| {
                let Value::Statement(statement) = value else {
                    unreachable!()
                };
                let StmtKind::Expr(value) = statement.kind else {
                    unreachable!()
                };
                Value::Expression(value)
            },
        );
        self.values.push(statement);
    }
    fn return_statement(&mut self, base: usize, span: vybe_ast::Span) {
        let children = self.collect(base);
        if children.is_empty() {
            self.push(Value::Statement(Box::new(Statement::with_span(
                StmtKind::Return(None),
                span,
            ))));
            return;
        }
        let value = if children.len() == 1 {
            children[0]
        } else {
            self.values.list(
                children,
                |values| {
                    Value::Expression(Expression::new(ExprKind::Tuple(
                        values.into_iter().map(expression).collect(),
                    )))
                },
                |value| {
                    let ExprKind::Tuple(values) = expression(value).kind else {
                        unreachable!()
                    };
                    values.into_iter().map(Value::Expression).collect()
                },
            )
        };
        let statement = self.values.unary(
            value,
            |value| {
                Value::Statement(Box::new(Statement::with_span(
                    StmtKind::Return(Some(expression(value))),
                    span,
                )))
            },
            |value| {
                let Value::Statement(statement) = value else {
                    unreachable!()
                };
                let StmtKind::Return(Some(value)) = statement.kind else {
                    unreachable!()
                };
                Value::Expression(value)
            },
        );
        self.values.push(statement);
    }
    fn jump_statement(&mut self, frame: Frame, span: vybe_ast::Span) {
        let children = self.collect(frame.base);
        assert_eq!(children.len(), 1);
        self.set_span(children[0], Default::default());
        let jump = frame.rule == Rule::goto_statement;
        let statement = self.values.unary(
            children[0],
            |value| {
                let ExprKind::Ident(name) = expression(value).kind else {
                    unreachable!()
                };
                let kind = if jump {
                    StmtKind::GoTo(name)
                } else {
                    StmtKind::Label(name)
                };
                Value::Statement(Box::new(Statement::with_span(kind, span)))
            },
            |value| {
                let Value::Statement(statement) = value else {
                    unreachable!()
                };
                let name = match statement.kind {
                    StmtKind::GoTo(name) | StmtKind::Label(name) => name,
                    _ => unreachable!(),
                };
                Value::Expression(Expression::new(ExprKind::Ident(name)))
            },
        );
        self.values.push(statement);
    }
    fn local_var(&mut self, frame: Frame, span: vybe_ast::Span) {
        let mut children = self.collect(frame.base);
        // Preserve the current walker's policy: extra initializer expressions
        // are parsed but omitted from this raw AST. Their live arena values
        // remain available to rollback, and are dropped at owned-root transfer.
        children.truncate(frame.names * 2);
        for name in &children[..frame.names] {
            self.set_span(*name, Default::default());
        }
        let statement = self.values.list(
            children,
            |values| {
                let mut values = values.into_iter();
                let names: Vec<_> = values
                    .by_ref()
                    .take(frame.names)
                    .map(|value| {
                        let ExprKind::Ident(name) = expression(value).kind else {
                            unreachable!()
                        };
                        name
                    })
                    .collect();
                let mut initializers = values.map(expression);
                let declarations = names
                    .into_iter()
                    .map(|name| VarDeclarator {
                        pattern: BindingPattern::Ident(name),
                        type_hint: None,
                        init: initializers.next(),
                        array_bounds: None,
                        with_events: false,
                    })
                    .collect();
                Value::Statement(Box::new(Statement::with_span(
                    StmtKind::VarDecl {
                        declarations,
                        kind: VarDeclKind::Let,
                    },
                    span,
                )))
            },
            |value| {
                let Value::Statement(statement) = value else {
                    unreachable!()
                };
                let StmtKind::VarDecl { declarations, .. } = statement.kind else {
                    unreachable!()
                };
                let mut values = Vec::with_capacity(declarations.len() * 2);
                let mut initializers = Vec::new();
                for declaration in declarations {
                    let BindingPattern::Ident(name) = declaration.pattern else {
                        unreachable!()
                    };
                    values.push(Value::Expression(Expression::new(ExprKind::Ident(name))));
                    if let Some(value) = declaration.init {
                        initializers.push(Value::Expression(value));
                    }
                }
                values.extend(initializers);
                values
            },
        );
        self.values.push(statement);
    }
    fn function(&mut self, frame: Frame, span: vybe_ast::Span) {
        if frame.method {
            self.method_function(frame, span);
            return;
        }
        let children = self.collect(frame.base);
        let declaration = matches!(frame.rule, Rule::local_function | Rule::function_decl);
        if declaration {
            self.set_span(children[0], Default::default());
        }
        let function = self.values.list(
            children,
            |values| {
                let mut values = values.into_iter();
                let name = if declaration {
                    Some(expression(values.next().unwrap()))
                } else {
                    None
                };
                let mut params = Vec::new();
                let first = values.next().unwrap();
                let body = match first {
                    Value::Params(value) => {
                        params = value;
                        let Value::Body(body) = values.next().unwrap() else {
                            unreachable!()
                        };
                        body
                    }
                    Value::Body(body) => body,
                    _ => unreachable!(),
                };
                assert!(values.next().is_none());
                if let Some(target) = name {
                    let ExprKind::Ident(name) = target.kind else {
                        return Value::Statement(Box::new(Statement::with_span(
                            StmtKind::Assign {
                                targets: vec![target],
                                value: Expression::new(ExprKind::Lambda {
                                    params,
                                    body: LambdaBody::Block(body),
                                    is_async: false,
                                    captures: Vec::new(),
                                }),
                                by_ref: false,
                            },
                            span,
                        )));
                    };
                    Value::Statement(Box::new(Statement::with_span(
                        StmtKind::FunctionDecl {
                            name,
                            params,
                            return_type: None,
                            body,
                            modifiers: Default::default(),
                            handles: Vec::new(),
                            is_async: false,
                            is_generator: false,
                            is_sub: false,
                        },
                        span,
                    )))
                } else {
                    Value::Expression(Expression::new(ExprKind::Lambda {
                        params,
                        body: LambdaBody::Block(body),
                        is_async: false,
                        captures: Vec::new(),
                    }))
                }
            },
            split_function,
        );
        self.values.push(function);
    }
    fn method_function(&mut self, frame: Frame, span: vybe_ast::Span) {
        let mut children = self.collect(frame.base);
        assert!(matches!(children.len(), 2 | 3));
        if children.len() == 2 {
            let params = self.values.alloc(Value::Params(Vec::new()));
            children.insert(1, params);
        }
        let Value::Expression(target) = self.values.get(children[0]) else {
            unreachable!()
        };
        let preserve = preserve_method_field(target);
        self.values.edit(
            children[1],
            |value| {
                let Value::Params(params) = value else {
                    unreachable!()
                };
                let inserted = !params.first().is_some_and(|param| param.name == "self");
                if inserted {
                    params.insert(
                        0,
                        Param {
                            name: "self".into(),
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        },
                    );
                }
                Metadata::InsertedSelf(inserted)
            },
            |value, previous| {
                let (Value::Params(params), Metadata::InsertedSelf(inserted)) = (value, previous)
                else {
                    unreachable!()
                };
                if inserted {
                    params.remove(0);
                }
            },
        );
        let function = self.values.list(
            children,
            |values| {
                let mut values = values.into_iter();
                let target = expression(values.next().unwrap());
                let Value::Params(params) = values.next().unwrap() else {
                    unreachable!()
                };
                let Value::Body(body) = values.next().unwrap() else {
                    unreachable!()
                };
                Value::Statement(Box::new(Statement::with_span(
                    StmtKind::Block(vec![
                        preserve,
                        Statement::new(StmtKind::Assign {
                            targets: vec![target],
                            value: Expression::new(ExprKind::Lambda {
                                params,
                                body: LambdaBody::Block(body),
                                is_async: false,
                                captures: Vec::new(),
                            }),
                            by_ref: false,
                        }),
                    ]),
                    span,
                )))
            },
            |value| {
                let Value::Statement(statement) = value else {
                    unreachable!()
                };
                let StmtKind::Block(mut statements) = statement.kind else {
                    unreachable!()
                };
                assert_eq!(statements.len(), 2);
                let StmtKind::Assign {
                    mut targets, value, ..
                } = statements.pop().unwrap().kind
                else {
                    unreachable!()
                };
                let ExprKind::Lambda {
                    params,
                    body: LambdaBody::Block(body),
                    ..
                } = value.kind
                else {
                    unreachable!()
                };
                vec![
                    Value::Expression(targets.pop().unwrap()),
                    Value::Params(params),
                    Value::Body(body),
                ]
            },
        );
        self.values.push(function);
    }
    fn unary(&mut self, operand: Id, operator: Operator) -> Id {
        if matches!(operator, Operator::Length) {
            self.values.unary(
                operand,
                |value| {
                    Value::Expression(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Ident("__lua_len".into()))),
                        args: vec![Argument::positional(expression(value))],
                        optional: false,
                    }))
                },
                |value| {
                    let ExprKind::Call { mut args, .. } = expression(value).kind else {
                        unreachable!()
                    };
                    Value::Expression(args.pop().unwrap().value)
                },
            )
        } else {
            let Operator::Unary(op) = operator else {
                unreachable!()
            };
            self.values.unary(
                operand,
                |value| {
                    Value::Expression(Expression::new(ExprKind::Unary {
                        op,
                        expr: Box::new(expression(value)),
                    }))
                },
                |value| {
                    let ExprKind::Unary { expr, .. } = expression(value).kind else {
                        unreachable!()
                    };
                    Value::Expression(*expr)
                },
            )
        }
    }
    fn parenthesized_call(&mut self, id: Id) -> Id {
        let Value::Expression(value) = self.values.get(id) else {
            unreachable!()
        };
        if !matches!(value.kind, ExprKind::Call { .. })
            || matches!(&value.kind, ExprKind::Call { callee, .. } if matches!(&callee.kind, ExprKind::Ident(name) if name == "__lua_first"))
        {
            return id;
        }
        self.values.unary(
            id,
            |value| {
                Value::Expression(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Ident("__lua_first".into()))),
                    args: vec![Argument::positional(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Ident(
                            "__lua_multi_row".into(),
                        ))),
                        args: vec![Argument::positional(expression(value))],
                        optional: false,
                    }))],
                    optional: false,
                }))
            },
            |value| {
                let ExprKind::Call { mut args, .. } = expression(value).kind else {
                    unreachable!()
                };
                let ExprKind::Call { mut args, .. } = args.pop().unwrap().value.kind else {
                    unreachable!()
                };
                Value::Expression(args.pop().unwrap().value)
            },
        )
    }
    fn call_chain(&mut self, base: usize, raw: &str) {
        let mut children = Vec::new();
        while self.values.values().len() > base {
            children.push(self.values.pop().unwrap());
        }
        children.reverse();
        let object = self.values.pop().unwrap();
        let root = match raw.as_bytes()[0] {
            b'[' => self.values.binary(
                object,
                children[0],
                |object, index| {
                    Value::Expression(Expression::new(ExprKind::Index {
                        object: Box::new(expression(object)),
                        index: Box::new(expression(index)),
                        null_safe: false,
                    }))
                },
                |value| {
                    let ExprKind::Index { object, index, .. } = expression(value).kind else {
                        unreachable!()
                    };
                    (Value::Expression(*object), Value::Expression(*index))
                },
            ),
            b'.' => {
                self.set_span(children[0], Default::default());
                self.values.binary(
                    object,
                    children[0],
                    |object, field| {
                        let ExprKind::Ident(field) = expression(field).kind else {
                            unreachable!()
                        };
                        Value::Expression(Expression::new(ExprKind::Member {
                            object: Box::new(expression(object)),
                            field,
                            null_safe: false,
                        }))
                    },
                    |value| {
                        let ExprKind::Member { object, field, .. } = expression(value).kind else {
                            unreachable!()
                        };
                        (
                            Value::Expression(*object),
                            Value::Expression(Expression::new(ExprKind::Ident(field))),
                        )
                    },
                )
            }
            b'(' | b':' => {
                let method = raw.starts_with(':');
                if method {
                    self.set_span(children[0], Default::default());
                }
                if !method {
                    let rename = match self.values.get(object) {
                        Value::Expression(Expression {
                            kind: ExprKind::Ident(name),
                            ..
                        }) => match name.as_str() {
                            "print" => Some("__lua_print"),
                            "tostring" => Some("__lua_tostring"),
                            "error" => Some("__lua_error"),
                            _ => None,
                        },
                        _ => None,
                    };
                    if let Some(rename) = rename {
                        self.values.edit(
                            object,
                            |value| {
                                let Value::Expression(Expression {
                                    kind: ExprKind::Ident(name),
                                    span,
                                }) = value
                                else {
                                    unreachable!()
                                };
                                Metadata::Ident {
                                    name: std::mem::replace(name, rename.into()),
                                    span: std::mem::take(span),
                                }
                            },
                            |value, previous| {
                                let (Value::Expression(value), Metadata::Ident { name, span }) =
                                    (value, previous)
                                else {
                                    unreachable!()
                                };
                                value.kind = ExprKind::Ident(name);
                                value.span = span;
                            },
                        );
                    }
                }
                children.insert(0, object);
                self.values.list(
                    children,
                    |values| {
                        let mut values = values.into_iter().map(expression);
                        let object = values.next().unwrap();
                        let (callee, mut args) = if method {
                            let ExprKind::Ident(field) = values.next().unwrap().kind else {
                                unreachable!()
                            };
                            (
                                Expression::new(ExprKind::Ident("__lua_method_call".into())),
                                vec![
                                    Argument::positional(object),
                                    Argument::positional(Expression::new(ExprKind::Lit(
                                        Literal::Str(field),
                                    ))),
                                ],
                            )
                        } else {
                            (object, Vec::new())
                        };
                        args.extend(values.map(Argument::positional));
                        Value::Expression(Expression::new(ExprKind::Call {
                            callee: Box::new(callee),
                            args,
                            optional: false,
                        }))
                    },
                    if method { split_method } else { split_call },
                )
            }
            _ => unreachable!(),
        };
        self.values.push(root);
    }
}
// The legacy method sugar repeats the table path in its generated checks.
// Materialize each required occurrence from scalar name components instead of
// cloning an already-owned AST subtree. The source target remains exclusively
// owned by the assignment and is recovered by the composition inverse.
fn preserve_method_field(target: &Expression) -> Statement {
    let ExprKind::Index {
        object,
        index: field_expr,
        ..
    } = &target.kind
    else {
        unreachable!()
    };
    let ExprKind::Lit(Literal::Str(field)) = &field_expr.kind else {
        unreachable!()
    };
    let mut suffixes = Vec::new();
    let mut base = object.as_ref();
    while let ExprKind::Index { object, index, .. } = &base.kind {
        let ExprKind::Lit(Literal::Str(name)) = &index.kind else {
            unreachable!()
        };
        suffixes.push(name.as_str());
        base = object;
    }
    let ExprKind::Ident(base) = &base.kind else {
        unreachable!()
    };
    fn key(name: &str) -> Expression {
        Expression::new(ExprKind::Lit(Literal::Str(name.into())))
    }
    fn index(object: Expression, index: Expression) -> Expression {
        Expression::new(ExprKind::Index {
            object: Box::new(object),
            index: Box::new(index),
            null_safe: false,
        })
    }
    fn binary(op: BinOp, left: Expression, right: Expression) -> Expression {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        })
    }
    fn unary(op: UnaryOp, expr: Expression) -> Expression {
        Expression::new(ExprKind::Unary {
            op,
            expr: Box::new(expr),
        })
    }
    let table = || {
        let mut table = Expression::new(ExprKind::Ident(base.to_owned()));
        for name in suffixes.iter().rev() {
            table = index(table, key(name));
        }
        table
    };
    let slot = || index(table(), key(field));
    let data_table = || index(table(), key("__lua_data"));
    let null = || Expression::new(ExprKind::Lit(Literal::Null));
    let cond = binary(
        BinOp::And,
        unary(
            UnaryOp::Not,
            binary(BinOp::Eq, unary(UnaryOp::Typeof, slot()), key("function")),
        ),
        binary(BinOp::NotEq, slot(), null()),
    );
    Statement::new(StmtKind::If {
        cond,
        then_body: vec![
            Statement::new(StmtKind::If {
                cond: binary(BinOp::Eq, data_table(), null()),
                then_body: vec![Statement::new(StmtKind::Assign {
                    targets: vec![data_table()],
                    value: Expression::new(ExprKind::Array(Vec::new())),
                    by_ref: false,
                })],
                elifs: Vec::new(),
                else_body: None,
            }),
            Statement::new(StmtKind::Assign {
                targets: vec![index(data_table(), key(field))],
                value: slot(),
                by_ref: false,
            }),
        ],
        elifs: Vec::new(),
        else_body: None,
    })
}
fn split_function(value: Value) -> Vec<Value> {
    let (name, params, body) = match value {
        Value::Expression(Expression {
            kind:
                ExprKind::Lambda {
                    params,
                    body: LambdaBody::Block(body),
                    ..
                },
            ..
        }) => (None, params, body),
        Value::Statement(statement) => match statement.kind {
            StmtKind::FunctionDecl {
                name, params, body, ..
            } => (Some(Expression::new(ExprKind::Ident(name))), params, body),
            StmtKind::Assign {
                mut targets, value, ..
            } => {
                assert_eq!(targets.len(), 1);
                let ExprKind::Lambda {
                    params,
                    body: LambdaBody::Block(body),
                    ..
                } = value.kind
                else {
                    unreachable!()
                };
                (Some(targets.pop().unwrap()), params, body)
            }
            _ => unreachable!(),
        },
        _ => unreachable!(),
    };
    let mut values = Vec::new();
    if let Some(name) = name {
        values.push(Value::Expression(name));
    }
    if !params.is_empty() {
        values.push(Value::Params(params));
    }
    values.push(Value::Body(body));
    values
}
fn split_single_rhs(value: Value) -> Value {
    Value::Expressions(vec![expression(value)])
}
fn split_multiple_rhs(value: Value) -> Value {
    let ExprKind::Array(fields) = expression(value).kind else {
        unreachable!()
    };
    Value::Expressions(fields.into_iter().map(|field| field.value).collect())
}
fn split_call(value: Value) -> Vec<Value> {
    let ExprKind::Call { callee, args, .. } = expression(value).kind else {
        unreachable!()
    };
    let mut values = vec![Value::Expression(*callee)];
    values.extend(args.into_iter().map(|arg| Value::Expression(arg.value)));
    values
}
fn split_method(value: Value) -> Vec<Value> {
    let ExprKind::Call { args, .. } = expression(value).kind else {
        unreachable!()
    };
    let mut args = args.into_iter();
    let object = args.next().unwrap().value;
    let ExprKind::Lit(Literal::Str(field)) = args.next().unwrap().value.kind else {
        unreachable!()
    };
    let mut values = vec![
        Value::Expression(object),
        Value::Expression(Expression::new(ExprKind::Ident(field))),
    ];
    values.extend(args.map(|arg| Value::Expression(arg.value)));
    values
}
fn lua_tier(op: BinOp) -> Option<u8> {
    Some(match op {
        BinOp::Or => 1,
        BinOp::And => 2,
        BinOp::Lt | BinOp::Gt | BinOp::LtEq | BinOp::GtEq | BinOp::Eq | BinOp::NotEq => 3,
        BinOp::BitOr => 4,
        BinOp::BitXor => 5,
        BinOp::BitAnd => 6,
        BinOp::Shl | BinOp::Shr => 7,
        BinOp::Concat => 8,
        BinOp::Add | BinOp::Sub => 9,
        BinOp::Mul | BinOp::Div | BinOp::FloorDiv | BinOp::Mod => 10,
        _ => return None,
    })
}
fn binop(text: &str) -> BinOp {
    match text {
        "and" => BinOp::And,
        "or" => BinOp::Or,
        "<" => BinOp::Lt,
        ">" => BinOp::Gt,
        "<=" => BinOp::LtEq,
        ">=" => BinOp::GtEq,
        "~=" => BinOp::NotEq,
        "==" => BinOp::Eq,
        ".." => BinOp::Concat,
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "//" => BinOp::FloorDiv,
        "%" => BinOp::Mod,
        "^" => BinOp::Pow,
        "&" => BinOp::BitAnd,
        "|" => BinOp::BitOr,
        "~" => BinOp::BitXor,
        "<<" => BinOp::Shl,
        ">>" => BinOp::Shr,
        _ => unreachable!("unknown Lua operator"),
    }
}
impl Builder for Ast<'_> {
    fn try_pratt_boundary(&mut self, _: usize, offset: usize) -> Result<(), BuildError> {
        let id = *self.values.values().last().unwrap();
        let Value::Expression(value) = self.values.get(id) else {
            unreachable!()
        };
        let mut span = value.span;
        let end = self.span(Span {
            start: offset,
            end: offset,
        });
        span.end_line = end.end_line;
        span.end_col = end.end_col;
        self.undo.push(Undo::Arena(self.values.checkpoint()));
        self.set_span(id, span);
        Ok(())
    }
    fn supports_pratt(&self, id: usize) -> bool {
        self.pratt && rule(id) == Rule::expr
    }
    fn try_reduce(
        &mut self,
        _: usize,
        fixity: grameo::pratt::Fixity,
        span: Span,
    ) -> Result<(), BuildError> {
        use grameo::pratt::Fixity;
        self.undo.push(Undo::Arena(self.values.checkpoint()));
        let raw = self.source[span.start..span.end].trim();
        let right = self.values.pop().unwrap();
        let Value::Expression(right_value) = self.values.get(right) else {
            unreachable!()
        };
        let end_line = right_value.span.end_line;
        let end_col = right_value.span.end_col;
        let (root, start_line, start_col) = if fixity == Fixity::Prefix {
            let operator = match raw {
                "not" => Operator::Unary(UnaryOp::Not),
                "-" => Operator::Unary(UnaryOp::Neg),
                "~" => Operator::Unary(UnaryOp::BitNot),
                "#" => Operator::Length,
                _ => unreachable!("unknown Lua prefix operator"),
            };
            let start = self.span(span);
            (
                self.unary(right, operator),
                start.start_line,
                start.start_col,
            )
        } else {
            assert!(matches!(fixity, Fixity::InfixLeft | Fixity::InfixRight));
            let left = self.values.pop().unwrap();
            let Value::Expression(left_value) = self.values.get(left) else {
                unreachable!()
            };
            let start_line = left_value.span.start_line;
            let start_col = left_value.span.start_col;
            let op = binop(raw);
            if fixity == Fixity::InfixLeft {
                self.flat_chain_child(left, op);
            } else if op == BinOp::Concat {
                self.flat_chain_child(right, op);
            }
            (self.binary(left, right, op), start_line, start_col)
        };
        self.set_span(
            root,
            vybe_ast::Span {
                start_line,
                start_col,
                end_line,
                end_col,
            },
        );
        self.values.push(root);
        if fixity != Fixity::Prefix
            && let Some(tier) = lua_tier(binop(raw))
        {
            let context = self.frames.last().unwrap().context;
            assert!(self.pratt_origins.insert(root, (context, tier)).is_none());
            self.undo.push(Undo::RemovePrattOrigin(root));
        }
        Ok(())
    }
    fn checkpoint(&self) -> usize {
        self.undo.len()
    }
    fn rollback(&mut self, mark: usize) {
        while self.undo.len() > mark {
            match self.undo.pop().unwrap() {
                Undo::Arena(mark) => self.values.rollback(mark),
                Undo::PopFrame => {
                    self.frames.pop().unwrap();
                }
                Undo::RestoreFrame(frame) => self.frames.push(frame),
                Undo::RestoreFrameState { index, frame } => self.frames[index] = frame,
                Undo::RemovePrattOrigin(id) => {
                    self.pratt_origins.remove(&id).unwrap();
                }
            }
        }
    }
    fn begin(&mut self, id: usize, _: usize) {
        let rule = rule(id);
        if wrapper(rule) {
            self.frames.push(Frame {
                rule,
                base: self.values.values().len(),
                field: None,
                names: 0,
                method: false,
                context: self.undo.len(),
            });
            self.undo.push(Undo::PopFrame);
        }
    }
    fn finish(&mut self, _: usize, _: Span, _: &str) {}
    fn try_finish(&mut self, id: usize, span: Span, _: &str) -> Result<(), BuildError> {
        let rule = rule(id);
        let raw = self.source[span.start..span.end].trim();
        match rule {
            Rule::name => {
                self.mark_field(FieldKind::Named);
                self.mark_local_name();
            }
            Rule::lbracket => self.mark_field(FieldKind::Computed),
            Rule::ELLIPSIS => self.mark_field(FieldKind::Spread),
            _ => {}
        }
        if rule == Rule::break_statement {
            self.undo.push(Undo::Arena(self.values.checkpoint()));
            self.push(Value::Statement(Box::new(Statement::with_span(
                StmtKind::Break(BreakTarget::Implicit),
                self.span(span),
            ))));
            return Ok(());
        }
        if wrapper(rule) {
            let frame = self.frames.pop().unwrap();
            assert_eq!(frame.rule, rule);
            self.undo.push(Undo::RestoreFrame(frame));
            self.undo.push(Undo::Arena(self.values.checkpoint()));
            match rule {
                Rule::chunk => {
                    let children = self.collect(frame.base);
                    assert_eq!(children.len(), 1);
                    let module = self.values.unary(
                        children[0],
                        |value| {
                            let Value::Body(body) = value else {
                                unreachable!()
                            };
                            Value::Module(Box::new(Module {
                                canon: Default::default(),
                                name: "main".into(),
                                language: Lang::Lua,
                                body,
                                imports: Vec::new(),
                                directives: Directives {
                                    method_receiver: Some(MethodReceiver::CallSite),
                                    receiver_binding: Some(ReceiverBinding::UniversalParameter),
                                    type_resolution: Some(TypeResolution::Hybrid),
                                    operator_dispatch: Some(OperatorDispatch::RuntimeProtocol),
                                    ..Default::default()
                                },
                            }))
                        },
                        |value| {
                            let Value::Module(module) = value else {
                                unreachable!()
                            };
                            Value::Body(module.body)
                        },
                    );
                    self.values.push(module);
                    return Ok(());
                }
                Rule::func_name => {
                    let children = self.collect(frame.base);
                    let mut children = children.into_iter();
                    let mut target = children.next().unwrap();
                    self.set_span(target, Default::default());
                    for field in children {
                        target = self.values.binary(
                            target,
                            field,
                            |object, index| {
                                Value::Expression(Expression::new(ExprKind::Index {
                                    object: Box::new(expression(object)),
                                    index: Box::new(expression(index)),
                                    null_safe: false,
                                }))
                            },
                            |value| {
                                let ExprKind::Index { object, index, .. } = expression(value).kind
                                else {
                                    unreachable!()
                                };
                                (Value::Expression(*object), Value::Expression(*index))
                            },
                        );
                    }
                    self.values.push(target);
                    return Ok(());
                }
                Rule::func_name_suffix => {
                    if raw.starts_with(':') {
                        let index = self
                            .frames
                            .iter()
                            .rposition(|frame| frame.rule == Rule::function_decl)
                            .unwrap();
                        let frame = self.frames[index];
                        self.undo.push(Undo::RestoreFrameState { index, frame });
                        self.frames[index].method = true;
                    }
                    let children = self.collect(frame.base);
                    assert_eq!(children.len(), 1);
                    let name = children[0];
                    self.set_span(name, Default::default());
                    let field = self.values.unary(
                        name,
                        |value| {
                            let ExprKind::Ident(name) = expression(value).kind else {
                                unreachable!()
                            };
                            Value::Expression(Expression::new(ExprKind::Lit(Literal::Str(name))))
                        },
                        |value| {
                            let ExprKind::Lit(Literal::Str(name)) = expression(value).kind else {
                                unreachable!()
                            };
                            Value::Expression(Expression::new(ExprKind::Ident(name)))
                        },
                    );
                    self.values.push(field);
                    return Ok(());
                }
                Rule::assign_lhs | Rule::assign_rhs => {
                    self.expression_list(frame.base);
                    return Ok(());
                }
                Rule::for_varlist => {
                    self.for_names(frame.base);
                    return Ok(());
                }
                Rule::for_numeric => {
                    self.numeric_for(frame, self.span(span));
                    return Ok(());
                }
                Rule::for_generic => {
                    self.generic_for(frame, self.span(span));
                    return Ok(());
                }
                Rule::assign_stmt => {
                    self.assignment(frame, self.span(span));
                    return Ok(());
                }
                Rule::do_statement | Rule::while_statement | Rule::repeat_statement => {
                    self.control_statement(frame, self.span(span));
                    return Ok(());
                }
                Rule::if_statement => {
                    self.conditional(frame, self.span(span));
                    return Ok(());
                }
                Rule::param => {
                    self.parameter(frame, raw == "...");
                    return Ok(());
                }
                Rule::param_list => {
                    self.parameters(frame.base);
                    return Ok(());
                }
                Rule::func_body
                | Rule::body_block
                | Rule::then_block
                | Rule::else_block
                | Rule::block => {
                    self.body(frame.base);
                    return Ok(());
                }
                Rule::func_expr | Rule::local_function | Rule::function_decl => {
                    self.function(frame, self.span(span));
                    return Ok(());
                }
                Rule::return_statement => {
                    self.return_statement(frame.base, self.span(span));
                    return Ok(());
                }
                Rule::local_var => {
                    self.local_var(frame, self.span(span));
                    return Ok(());
                }
                Rule::label_statement | Rule::goto_statement => {
                    self.jump_statement(frame, self.span(span));
                    return Ok(());
                }
                _ => {}
            }
            if rule == Rule::table_field {
                self.table_field(frame);
                return Ok(());
            }
            if rule == Rule::table_constructor {
                self.table_constructor(frame.base);
                return Ok(());
            }
            if rule == Rule::call_chain {
                self.call_chain(frame.base, raw);
                return Ok(());
            }
            if rule == Rule::long_string && self.values.values().len() == frame.base {
                self.push(Value::Expression(Expression::new(ExprKind::Lit(
                    Literal::Str(literals::strip_long_brackets(raw).into()),
                ))));
            }
            if rule == Rule::primary && self.values.values().len() == frame.base {
                let kind = match raw {
                    "true" => ExprKind::Lit(Literal::Bool(true)),
                    "false" => ExprKind::Lit(Literal::Bool(false)),
                    "nil" => ExprKind::Lit(Literal::Null),
                    "..." => ExprKind::Spread(Box::new(Expression::null())),
                    _ => unreachable!("unknown Lua primary"),
                };
                self.push(Value::Expression(Expression::new(kind)));
            }
            if self.values.values().len() == frame.base + 1 {
                let mut root = *self.values.values().last().unwrap();
                if !matches!(
                    rule,
                    Rule::call_expression | Rule::postfix | Rule::long_string
                ) && raw.starts_with('(')
                    && raw.ends_with(')')
                {
                    self.values.pop();
                    root = self.parenthesized_call(root);
                    self.values.push(root);
                }
                self.set_span(root, self.span(span));
                self.expression_statement(rule);
                return Ok(());
            }
            let mut items = Vec::new();
            while self.values.values().len() > frame.base {
                items.push(self.values.pop().unwrap());
            }
            items.reverse();
            let mut root = if matches!(self.values.get(items[0]), Value::Operator(_)) {
                let Value::Operator(operator) = self.values.get(items[0]) else {
                    unreachable!()
                };
                self.unary(items[1], *operator)
            } else if matches!(rule, Rule::pow_expr | Rule::concat_expr) && items.len() > 1 {
                let op = if rule == Rule::pow_expr {
                    BinOp::Pow
                } else {
                    BinOp::Concat
                };
                let mut root = *items.last().unwrap();
                for left in items[..items.len() - 1].iter().step_by(2).rev() {
                    root = self.binary(*left, root, op);
                }
                root
            } else {
                let mut root = items[0];
                for pair in items[1..].chunks_exact(2) {
                    let Value::Operator(Operator::Binary(op)) = self.values.get(pair[0]) else {
                        unreachable!()
                    };
                    root = self.binary(root, pair[1], *op);
                }
                root
            };
            if !matches!(
                rule,
                Rule::call_expression | Rule::postfix | Rule::long_string
            ) && raw.starts_with('(')
                && raw.ends_with(')')
                && items.len() == 1
            {
                root = self.parenthesized_call(root);
            }
            self.set_span(root, self.span(span));
            self.values.push(root);
            self.expression_statement(rule);
        } else if matches!(
            rule,
            Rule::name | Rule::number | Rule::string | Rule::long_string_body_0
        ) {
            let kind = match rule {
                Rule::name => ExprKind::Ident(raw.into()),
                Rule::number => {
                    let value = if raw.starts_with("0x") || raw.starts_with("0X") {
                        Literal::Int(i64::from_str_radix(&raw[2..], 16).unwrap_or(0))
                    } else if raw.contains(['.', 'e', 'E']) {
                        Literal::Float(raw.parse().unwrap_or(0.0))
                    } else {
                        Literal::Int(raw.parse().unwrap_or(0))
                    };
                    ExprKind::Lit(value)
                }
                Rule::string => ExprKind::Lit(Literal::Str(literals::unescape_lua_string(
                    &raw[1..raw.len() - 1],
                ))),
                Rule::long_string_body_0 => {
                    let body = &self.source[span.start..span.end];
                    let body = body
                        .strip_prefix("\r\n")
                        .or_else(|| body.strip_prefix('\n'))
                        .unwrap_or(body);
                    ExprKind::Lit(Literal::Str(body.into()))
                }
                _ => unreachable!(),
            };
            self.undo.push(Undo::Arena(self.values.checkpoint()));
            self.push(Value::Expression(Expression::with_span(
                kind,
                self.span(span),
            )));
        } else if matches!(
            rule,
            Rule::unop
                | Rule::pow_op
                | Rule::mul_op
                | Rule::additive_op
                | Rule::shift_op
                | Rule::compare_op
                | Rule::CONCAT
                | Rule::AMP
                | Rule::TILDE
                | Rule::PIPE
                | Rule::KW_AND
                | Rule::KW_OR
        ) {
            if rule == Rule::TILDE
                && self
                    .frames
                    .last()
                    .is_some_and(|frame| frame.rule == Rule::unary)
            {
                return Ok(());
            }
            self.undo.push(Undo::Arena(self.values.checkpoint()));
            let operator = if rule == Rule::unop {
                match raw {
                    "not" => Operator::Unary(UnaryOp::Not),
                    "-" => Operator::Unary(UnaryOp::Neg),
                    "~" => Operator::Unary(UnaryOp::BitNot),
                    "#" => Operator::Length,
                    _ => unreachable!(),
                }
            } else {
                Operator::Binary(binop(raw))
            };
            self.push(Value::Operator(operator));
        }
        Ok(())
    }
}
