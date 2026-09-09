//! `[System.Management.Automation.Language.Parser]::ParseInput(…)` — the AST
//! surface a PowerShell script inspects, TRANSLATED from the common AST.
//!
//! ⛔THERE IS ONE AST. The script believes it is running on PowerShell and
//! reading PowerShell's own parse tree; it is running on WASM, over the common
//! AST every frontend targets, and this module hands it EQUIVALENT information.
//! Nothing here re-parses and nothing here builds a second tree: `walker::parse`
//! is the one parse, and the nodes it produced are what these objects describe.
//!
//! ⛔WHERE NORMALIZING DROPPED THE SURFACE FORM, THE SPAN CARRIES IT. A single-
//! and a double-quoted string are both `Lit(Str)`, and `switch` normalizes to
//! `If` — the common AST is the LOWERED form, which is its job. Each node now
//! carries the source span the walker stamps, so the spelling is recoverable
//! from the text it points at: a leading `'` is `SingleQuoted`, a statement
//! whose text opens with `switch` is a `SwitchStatementAst`, and its flags are
//! its own `-Regex` / `-Wildcard` words. That is the translation, and it is why
//! no second representation is needed.
//!
//! The delivery is the shape `php/src/emitter/reflection_adapter.rs` describes:
//! the walker captures the metadata at parse time and the objects are built
//! from it, with members bound as named properties.

use vybe_ast::{
    Argument, ArrayElement, BindingPattern, ExprKind, Expression, LambdaBody, Literal, Module,
    Modifiers, ObjectProperty, Param, PassBy, Span, Statement, StmtKind, VarDeclKind,
    VarDeclarator,
};

/// One node of the surface tree: what the script sees.
struct AstNode {
    ty: String,
    span: Span,
    children: Vec<usize>,
    members: Vec<(String, Member)>,
}

enum Member {
    Node(usize),
    Nodes(Vec<usize>),
    Expr(Expression),
    Str(String),
    Int(i64),
    Bool(bool),
    Null,
}

/// The source text a span covers, and the offsets of its ends.
struct SourceMap<'a> {
    source: &'a str,
    /// Byte offset of the first character of each 1-based line.
    line_starts: Vec<usize>,
}

impl<'a> SourceMap<'a> {
    fn new(source: &'a str) -> Self {
        let mut line_starts = vec![0usize];
        for (index, byte) in source.bytes().enumerate() {
            if byte == b'\n' {
                line_starts.push(index + 1);
            }
        }
        SourceMap {
            source,
            line_starts,
        }
    }

    fn offset(&self, line: u32, col: u32) -> usize {
        if line == 0 {
            return 0;
        }
        let base = *self
            .line_starts
            .get(line as usize - 1)
            .unwrap_or(&0);
        (base + col.saturating_sub(1) as usize).min(self.source.len())
    }

    fn text(&self, span: Span) -> &'a str {
        let start = self.offset(span.start_line, span.start_col);
        let end = self.offset(span.end_line, span.end_col).max(start);
        self.source.get(start..end).unwrap_or("")
    }
}

/// Build the object graph for `source`, already parsed into `module`.
///
/// The result is an expression evaluating to the root `ScriptBlockAst`.
pub fn ast_expr(module: &Module, source: &str) -> Expression {
    let map = SourceMap::new(source);
    let mut nodes: Vec<AstNode> = Vec::new();
    let mut roots = Vec::new();
    let whole_span = Span {
        start_line: 1,
        start_col: 1,
        end_line: map.line_starts.len() as u32,
        end_col: u32::MAX,
    };
    if let Some(index) = source_compound_node(source.trim(), whole_span, &mut nodes) {
        roots.push(index);
    } else {
        for stmt in &module.body {
            if is_synthetic_span(stmt.span) {
                continue;
            }
            if let Some(index) = walk_stmt(stmt, &map, &mut nodes) {
                roots.push(index);
            }
        }
    }
    let param_block = param_block_node(source, whole_span, &mut nodes);
    let end_block = push(
        &mut nodes,
        "NamedBlockAst",
        whole_span,
        roots.clone(),
        vec![("Statements".to_string(), Member::Nodes(roots.clone()))],
    );
    let mut root_children = Vec::new();
    root_children.extend(param_block);
    root_children.push(end_block);
    let root = push(
        &mut nodes,
        "ScriptBlockAst",
        whole_span,
        root_children,
        vec![
            ("ParamBlock".to_string(), param_block.map(Member::Node).unwrap_or(Member::Null)),
            ("BeginBlock".to_string(), Member::Null),
            ("ProcessBlock".to_string(), Member::Null),
            ("EndBlock".to_string(), Member::Node(end_block)),
        ],
    );
    // The root's extent is the whole script, whatever the last line's length.
    nodes[root].members.push((
        "__WholeText".to_string(),
        Member::Str(source.to_string()),
    ));
    emit_tree(&nodes, root, &map, source)
}

pub fn ast_expr_lossy_source(source: &str) -> Expression {
    let map = SourceMap::new(source);
    let mut nodes: Vec<AstNode> = Vec::new();
    let whole_span = Span {
        start_line: 1,
        start_col: 1,
        end_line: map.line_starts.len() as u32,
        end_col: u32::MAX,
    };
    let roots = if let Some(index) = source_compound_node(source.trim(), whole_span, &mut nodes) {
        vec![index]
    } else {
        Vec::new()
    };
    let end_block = push(
        &mut nodes,
        "NamedBlockAst",
        whole_span,
        roots.clone(),
        vec![("Statements".to_string(), Member::Nodes(roots.clone()))],
    );
    let root = push(
        &mut nodes,
        "ScriptBlockAst",
        whole_span,
        vec![end_block],
        vec![
            ("ParamBlock".to_string(), Member::Null),
            ("BeginBlock".to_string(), Member::Null),
            ("ProcessBlock".to_string(), Member::Null),
            ("EndBlock".to_string(), Member::Node(end_block)),
        ],
    );
    nodes[root].members.push((
        "__WholeText".to_string(),
        Member::Str(source.to_string()),
    ));
    emit_tree(&nodes, root, &map, source)
}

fn push(
    nodes: &mut Vec<AstNode>,
    ty: &str,
    span: Span,
    children: Vec<usize>,
    members: Vec<(String, Member)>,
) -> usize {
    nodes.push(AstNode {
        ty: ty.to_string(),
        span,
        children,
        members,
    });
    nodes.len() - 1
}

fn is_synthetic_span(span: Span) -> bool {
    span.start_line == 0 && span.start_col == 0 && span.end_line == 0 && span.end_col == 0
}

fn statement_block_node(
    body: &[Statement],
    span: Span,
    map: &SourceMap,
    nodes: &mut Vec<AstNode>,
) -> usize {
    let statements: Vec<usize> = body
        .iter()
        .filter(|stmt| !is_synthetic_span(stmt.span))
        .filter_map(|stmt| walk_stmt(stmt, map, nodes))
        .collect();
    push(
        nodes,
        "StatementBlockAst",
        span,
        statements.clone(),
        vec![("Statements".to_string(), Member::Nodes(statements))],
    )
}

fn parameter_nodes(params: &[Param], span: Span, nodes: &mut Vec<AstNode>) -> Vec<usize> {
    params
        .iter()
        .filter(|param| param.name != "_")
        .map(|param| parameter_node(&param.name, span, Vec::new(), None, None, nodes))
        .collect()
}

fn parameter_node(
    name: &str,
    span: Span,
    attrs: Vec<usize>,
    default_value: Option<usize>,
    static_type: Option<String>,
    nodes: &mut Vec<AstNode>,
) -> usize {
    let variable = push(
        nodes,
        "VariableExpressionAst",
        span,
        Vec::new(),
        vec![("VariablePath".to_string(), Member::Expr(variable_path_expr(name)))],
    );
    let mut children = vec![variable];
    children.extend(attrs.iter().copied());
    children.extend(default_value);
    let static_type_expr = type_ref_expr(
        &static_type
            .as_deref()
            .and_then(type_accelerator)
            .unwrap_or_else(|| static_type.as_deref().unwrap_or("System.Object"))
            .to_string(),
        false,
    );
    push(
        nodes,
        "ParameterAst",
        span,
        children,
        vec![
            ("Name".to_string(), Member::Node(variable)),
            ("Attributes".to_string(), Member::Nodes(attrs)),
            ("DefaultValue".to_string(), default_value.map(Member::Node).unwrap_or(Member::Null)),
            ("StaticType".to_string(), Member::Expr(static_type_expr)),
        ],
    )
}

// ── translation ───────────────────────────────────────────────────────────

fn walk_stmt(stmt: &Statement, map: &SourceMap, nodes: &mut Vec<AstNode>) -> Option<usize> {
    let text = map.text(stmt.span).trim_start();
    match &stmt.kind {
        StmtKind::Assign { targets, value, .. } => {
            if let Some(assign) = source_assignment_node(text, stmt.span, nodes) {
                return Some(assign);
            }
            let left = targets.first().and_then(|t| walk_expr(t, map, nodes));
            let right = walk_expr(value, map, nodes);
            let left = typed_assignment_left(text, left, stmt.span, nodes);
            let mut children = Vec::new();
            children.extend(left);
            children.extend(right);
            let operator = assignment_operator(text);
            let members = vec![
                (
                    "Left".to_string(),
                    left.map(Member::Node).unwrap_or(Member::Null),
                ),
                (
                    "Right".to_string(),
                    right.map(Member::Node).unwrap_or(Member::Null),
                ),
                ("Operator".to_string(), Member::Str(operator_name(&operator))),
                ("__OperatorText".to_string(), Member::Str(operator)),
            ];
            Some(push(
                nodes,
                "AssignmentStatementAst",
                stmt.span,
                children,
                members,
            ))
        }
        StmtKind::Expr(expr) => {
            let text = map.text(stmt.span).trim();
            if text.starts_with('{') {
                return Some(push(
                    nodes,
                    "ScriptBlockExpressionAst",
                    stmt.span,
                    Vec::new(),
                    Vec::new(),
                ));
            }
            if let Some(type_name) = leading_bracketed_type(text) {
                return Some(type_constraint_node(&type_name, stmt.span, nodes));
            }
            if text.contains('|')
                || source_looks_like_command(text)
                || command_invocation_operator(text) != "Unknown"
                || has_redirection(text)
                || is_pipeline_expression_text(text)
            {
                return Some(pipeline_ast_node(text, stmt.span, nodes));
            }
            walk_expr(expr, map, nodes)
        }
        // ⛔`switch` NORMALIZES TO `If`. The lowering is right for execution and
        // wrong for inspection, and the span says which one the source spelled.
        StmtKind::If { .. } if source_switch_keyword(text).is_some() => {
            source_compound_node(text, stmt.span, nodes)
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            let mut children = Vec::new();
            let then_block = statement_block_node(then_body, stmt.span, map, nodes);
            children.push(then_block);
            let mut clauses = vec![tuple2_expr(Expression::null(), Expression::ident(&temp(then_block)))];
            for (_, body) in elifs {
                let block = statement_block_node(body, stmt.span, map, nodes);
                children.push(block);
                clauses.push(tuple2_expr(Expression::null(), Expression::ident(&temp(block))));
            }
            let else_clause = if let Some(body) = else_body {
                let block = statement_block_node(body, stmt.span, map, nodes);
                children.push(block);
                Member::Node(block)
            } else {
                Member::Null
            };
            Some(push(
                nodes,
                "IfStatementAst",
                stmt.span,
                children,
                vec![
                    ("Clauses".to_string(), Member::Expr(array_of(clauses))),
                    ("ElseClause".to_string(), else_clause),
                ],
            ))
        }
        StmtKind::While { .. } if text.starts_with("do") || text.starts_with(':') => {
            source_compound_node(text, stmt.span, nodes)
        }
        StmtKind::While { .. } => source_compound_node(text, stmt.span, nodes).or_else(|| {
            let children = child_statements(stmt, map, nodes);
            Some(push(
                nodes,
                "WhileStatementAst",
                stmt.span,
                children,
                Vec::new(),
            ))
        }),
        StmtKind::For { .. } => source_compound_node(text, stmt.span, nodes).or_else(|| {
            let children = child_statements(stmt, map, nodes);
            Some(push(
                nodes,
                "ForStatementAst",
                stmt.span,
                children,
                Vec::new(),
            ))
        }),
        StmtKind::ForIn { .. } if source_loop_keyword(text, "foreach").is_some() => {
            source_compound_node(text, stmt.span, nodes)
        }
        StmtKind::ForIn { .. } => {
            let children = child_statements(stmt, map, nodes);
            let var = foreach_variable(text);
            let mut members = Vec::new();
            if let Some(var) = var {
                let variable = push(
                    nodes,
                    "VariableExpressionAst",
                    stmt.span,
                    Vec::new(),
                    vec![("VariablePath".to_string(), Member::Expr(variable_path_expr(&var)))],
                );
                members.push(("Variable".to_string(), Member::Node(variable)));
            }
            Some(push(
                nodes,
                "ForEachStatementAst",
                stmt.span,
                children,
                members,
            ))
        }
        StmtKind::Try {
            body,
            catches,
            finally,
            ..
        } => {
            if starts_keyword(text.trim_start(), "try") {
                return Some(source_try_node(text, stmt.span, nodes));
            }
            let body_block = statement_block_node(body, stmt.span, map, nodes);
            let mut children = vec![body_block];
            let mut catch_nodes = Vec::new();
            for catch in catches {
                let catch_body = statement_block_node(&catch.body, stmt.span, map, nodes);
                let catch_node = push(
                    nodes,
                    "CatchClauseAst",
                    stmt.span,
                    vec![catch_body],
                    vec![
                        ("Body".to_string(), Member::Node(catch_body)),
                        ("IsCatchAll".to_string(), Member::Bool(catch.types.is_empty())),
                    ],
                );
                children.push(catch_node);
                catch_nodes.push(catch_node);
            }
            let finally_member = if let Some(finally) = finally {
                let block = statement_block_node(finally, stmt.span, map, nodes);
                children.push(block);
                Member::Node(block)
            } else {
                Member::Null
            };
            Some(push(
                nodes,
                "TryStatementAst",
                stmt.span,
                children,
                vec![
                    ("Body".to_string(), Member::Node(body_block)),
                    ("CatchClauses".to_string(), Member::Nodes(catch_nodes)),
                    ("Finally".to_string(), finally_member),
                ],
            ))
        }
        StmtKind::Throw { .. } => {
            return Some(source_throw_node(text, stmt.span, nodes));
        }
        StmtKind::Return(_) => {
            let children = child_statements(stmt, map, nodes);
            Some(push(
                nodes,
                "ReturnStatementAst",
                stmt.span,
                children,
                Vec::new(),
            ))
        }
        StmtKind::Break { .. } => Some(push(
            nodes,
            "BreakStatementAst",
            stmt.span,
            Vec::new(),
            Vec::new(),
        )),
        StmtKind::Continue { .. } => Some(push(
            nodes,
            "ContinueStatementAst",
            stmt.span,
            Vec::new(),
            Vec::new(),
        )),
        StmtKind::FunctionDecl { name, params, .. } => {
            let children = child_statements(stmt, map, nodes);
            let param_nodes = parameter_nodes(params, stmt.span, nodes);
            let mut all_children = children.clone();
            all_children.extend(param_nodes.iter().copied());
            let lower = text.trim_start().to_ascii_lowercase();
            let members = vec![
                ("Name".to_string(), Member::Str(name.clone())),
                ("IsFilter".to_string(), Member::Bool(lower.starts_with("filter"))),
                ("Parameters".to_string(), Member::Nodes(param_nodes)),
                ("Body".to_string(), Member::Expr(function_body_expr(text))),
            ];
            Some(push(
                nodes,
                "FunctionDefinitionAst",
                stmt.span,
                all_children,
                members,
            ))
        }
        StmtKind::VarDecl { declarations, .. } => {
            if let Some(assign) = source_assignment_node(text, stmt.span, nodes) {
                return Some(assign);
            }
            let mut children = Vec::new();
            if let Some(type_name) = leading_bracketed_type(text) {
                children.push(type_constraint_node(&type_name, stmt.span, nodes));
            }
            for decl in declarations {
                if let Some(init) = &decl.init {
                    children.extend(walk_expr(init, map, nodes));
                }
            }
            Some(push(
                nodes,
                "AssignmentStatementAst",
                stmt.span,
                children,
                Vec::new(),
            ))
        }
        StmtKind::Block(body) => {
            let children = body
                .iter()
                .filter_map(|s| walk_stmt(s, map, nodes))
                .collect();
            Some(push(
                nodes,
                "StatementBlockAst",
                stmt.span,
                children,
                Vec::new(),
            ))
        }
        _ => None,
    }
}

/// The statements a compound statement guards, walked as its children.
fn child_statements(stmt: &Statement, map: &SourceMap, nodes: &mut Vec<AstNode>) -> Vec<usize> {
    let mut out = Vec::new();
    let mut visit = |s: &Statement, nodes: &mut Vec<AstNode>| {
        if let Some(index) = walk_stmt(s, map, nodes) {
            out.push(index);
        }
    };
    match &stmt.kind {
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            for s in then_body {
                visit(s, nodes);
            }
            for (_, body) in elifs {
                for s in body {
                    visit(s, nodes);
                }
            }
            if let Some(else_body) = else_body {
                for s in else_body {
                    visit(s, nodes);
                }
            }
        }
        StmtKind::While { body, .. }
        | StmtKind::For { body, .. }
        | StmtKind::ForIn { body, .. }
        | StmtKind::FunctionDecl { body, .. } => {
            for s in body {
                visit(s, nodes);
            }
        }
        StmtKind::Try {
            body,
            catches,
            finally,
            ..
        } => {
            for s in body {
                visit(s, nodes);
            }
            for catch in catches {
                for s in &catch.body {
                    visit(s, nodes);
                }
            }
            if let Some(finally) = finally {
                for s in finally {
                    visit(s, nodes);
                }
            }
        }
        _ => {}
    }
    out
}

fn walk_expr(expr: &Expression, map: &SourceMap, nodes: &mut Vec<AstNode>) -> Option<usize> {
    let text = map.text(expr.span);
    if let Some(index) = source_static_member_ast(text, expr.span, nodes) {
        return Some(index);
    }
    if let Some(type_name) = source_bracketed_type(text) {
        return Some(type_expression_node(&type_name, expr.span, nodes));
    }
    if let Some(index) = source_member_ast(text, expr.span, nodes) {
        return Some(index);
    }
    match &expr.kind {
        ExprKind::Ident(name) => Some(push(
            nodes,
            "VariableExpressionAst",
            expr.span,
            Vec::new(),
            vec![(
                "VariablePath".to_string(),
                Member::Expr(variable_path_expr(name.trim_start_matches('$'))),
            )],
        )),
        // ⛔THE QUOTE IS NOT IN THE NODE. Both spellings are `Lit(Str)`; the
        // span's first character is what tells them apart, and PowerShell's
        // `StringConstantType` is exactly that distinction.
        ExprKind::Lit(Literal::Str(value)) => {
            let single = text.trim_start().starts_with('\'');
            let ty = if single {
                "StringConstantExpressionAst"
            } else if text.trim_start().starts_with('"') || text.trim_start().starts_with("@\"") {
                "ExpandableStringExpressionAst"
            } else {
                "StringConstantExpressionAst"
            };
            let kind = string_constant_kind(text, single);
            Some(push(
                nodes,
                ty,
                expr.span,
                Vec::new(),
                vec![
                    ("Value".to_string(), Member::Str(value.clone())),
                    ("StringConstantType".to_string(), Member::Str(kind)),
                    ("StaticType".to_string(), Member::Expr(type_ref_expr("System.String", false))),
                ],
            ))
        }
        ExprKind::Lit(Literal::Int(value)) => Some(push(
            nodes,
            "ConstantExpressionAst",
            expr.span,
            Vec::new(),
            vec![("Value".to_string(), Member::Int(*value))],
        )),
        ExprKind::Lit(Literal::Float(value)) => Some(push(
            nodes,
            "ConstantExpressionAst",
            expr.span,
            Vec::new(),
            vec![("Value".to_string(), Member::Str(value.to_string()))],
        )),
        ExprKind::Lit(Literal::Char(value)) => Some(push(
            nodes,
            "StringConstantExpressionAst",
            expr.span,
            Vec::new(),
            vec![
                ("Value".to_string(), Member::Str(value.to_string())),
                ("StringConstantType".to_string(), Member::Str("SingleQuoted".to_string())),
                ("StaticType".to_string(), Member::Expr(type_ref_expr("System.String", false))),
            ],
        )),
        ExprKind::Lit(Literal::Bool(value)) => Some(push(
            nodes,
            "ConstantExpressionAst",
            expr.span,
            Vec::new(),
            vec![("Value".to_string(), Member::Bool(*value))],
        )),
        ExprKind::Binary { left, right, .. } => {
            let a = walk_expr(left, map, nodes);
            let b = walk_expr(right, map, nodes);
            let children: Vec<usize> = a.into_iter().chain(b).collect();
            let members = vec![
                ("Operator".to_string(), Member::Str(binary_operator_name(text))),
                ("Left".to_string(), a.map(Member::Node).unwrap_or(Member::Null)),
                (
                    "Right".to_string(),
                    b.map(Member::Node).unwrap_or(Member::Null),
                ),
            ];
            Some(push(
                nodes,
                "BinaryExpressionAst",
                expr.span,
                children,
                members,
            ))
        }
        ExprKind::Cast { type_name, expr: inner } => {
            let child = walk_expr(inner, map, nodes);
            let ty = push(
                nodes,
                "TypeExpressionAst",
                expr.span,
                Vec::new(),
                vec![("TypeName".to_string(), Member::Expr(type_name_expr(type_name)))],
            );
            let children: Vec<usize> = std::iter::once(ty).chain(child).collect();
            Some(push(
                nodes,
                "ConvertExpressionAst",
                expr.span,
                children,
                vec![
                    ("Type".to_string(), Member::Node(ty)),
                    ("Child".to_string(), child.map(Member::Node).unwrap_or(Member::Null)),
                ],
            ))
        }
        ExprKind::Array(items) => {
            let children: Vec<usize> = items
                .iter()
                .filter_map(|item| walk_expr(&item.value, map, nodes))
                .collect();
            Some(push(
                nodes,
                "ArrayLiteralAst",
                expr.span,
                children.clone(),
                vec![("Elements".to_string(), Member::Nodes(children))],
            ))
        }
        ExprKind::Object(props) => {
            let mut children = Vec::new();
            let mut pairs = Vec::new();
            for prop in props {
                if let ObjectProperty::KeyValue { key, value } = prop {
                    if literal_text(key).is_some_and(|name| name.starts_with("__ps_")) {
                        continue;
                    }
                    let key_node = walk_expr(key, map, nodes);
                    let value_node = walk_expr(value, map, nodes);
                    children.extend(key_node);
                    children.extend(value_node);
                    if let (Some(k), Some(v)) = (key_node, value_node) {
                        pairs.push(tuple2_expr(Expression::ident(&temp(k)), Expression::ident(&temp(v))));
                    }
                }
            }
            Some(push(
                nodes,
                "HashtableAst",
                expr.span,
                children,
                vec![("KeyValuePairs".to_string(), Member::Expr(array_of(pairs)))],
            ))
        }
        ExprKind::Lambda { .. } => Some(push(
            nodes,
            "ScriptBlockExpressionAst",
            expr.span,
            Vec::new(),
            Vec::new(),
        )),
        ExprKind::Interpolation(parts) => {
            let nested: Vec<usize> = parts
                .iter()
                .filter_map(|part| match part {
                    vybe_ast::InterpolPart::Expr(expr) | vybe_ast::InterpolPart::Formatted(expr, _) => {
                        walk_expr(expr, map, nodes)
                    }
                    vybe_ast::InterpolPart::Text(_) => None,
                })
                .collect();
            Some(push(
                nodes,
                "ExpandableStringExpressionAst",
                expr.span,
                nested.clone(),
                vec![
                    ("Value".to_string(), Member::Str(text.trim_matches('"').to_string())),
                    ("StringConstantType".to_string(), Member::Str(string_constant_kind(text, false))),
                    ("NestedExpressions".to_string(), Member::Nodes(nested)),
                    ("StaticType".to_string(), Member::Expr(type_ref_expr("System.String", false))),
                ],
            ))
        }
        // ⛔A PIPELINE IS FOLDED INTO METHOD CALLS. Execution wants the fold and
        // inspection wants the stages; the span says which the source spelled,
        // and its own `|` marks the boundaries.
        ExprKind::Call { callee, args, .. } => {
            if text.trim().is_empty() && command_name(callee, text).is_empty() {
                return None;
            }
            let mut children = Vec::new();
            if let Some(index) = walk_expr(callee, map, nodes) {
                children.push(index);
            }
            for arg in args {
                if let Some(index) = walk_expr(&arg.value, map, nodes) {
                    children.push(index);
                }
            }
            let ty = if text.contains('|') {
                "PipelineAst"
            } else {
                "CommandAst"
            };
            let members = if ty == "PipelineAst" {
                let pipeline_elements = pipeline_element_nodes(text, expr.span, nodes);
                if !pipeline_elements.is_empty() {
                    children = pipeline_elements.clone();
                }
                vec![("PipelineElements".to_string(), Member::Nodes(children.clone()))]
            } else {
                let command_elements = command_element_nodes(text, expr.span, nodes);
                if !command_elements.is_empty() {
                    children = command_elements.clone();
                }
                vec![
                    ("CommandElements".to_string(), Member::Nodes(children.clone())),
                    ("CommandName".to_string(), Member::Str(command_name(callee, text))),
                ]
            };
            Some(push(nodes, ty, expr.span, children, members))
        }
        ExprKind::Member { object, field, .. } => {
            let target = source_static_type_from_member_text(text, expr.span, nodes)
                .or_else(|| walk_expr(object, map, nodes));
            let member = push(
                nodes,
                "StringConstantExpressionAst",
                expr.span,
                Vec::new(),
                vec![
                    ("Value".to_string(), Member::Str(field.clone())),
                    ("StringConstantType".to_string(), Member::Str("BareWord".to_string())),
                    ("StaticType".to_string(), Member::Expr(type_ref_expr("System.String", false))),
                ],
            );
            let children: Vec<usize> = target.into_iter().chain(std::iter::once(member)).collect();
            Some(push(
                nodes,
                "MemberExpressionAst",
                expr.span,
                children,
                vec![
                    ("Expression".to_string(), target.map(Member::Node).unwrap_or(Member::Null)),
                    ("Member".to_string(), Member::Node(member)),
                ],
            ))
        }
        ExprKind::Index { object, index, .. } => {
            let target = walk_expr(object, map, nodes);
            let idx = walk_expr(index, map, nodes);
            let children: Vec<usize> = target.into_iter().chain(idx).collect();
            Some(push(
                nodes,
                "IndexExpressionAst",
                expr.span,
                children,
                vec![
                    ("Target".to_string(), target.map(Member::Node).unwrap_or(Member::Null)),
                    ("Index".to_string(), idx.map(Member::Node).unwrap_or(Member::Null)),
                ],
            ))
        }
        ExprKind::Unary { expr: inner, .. } => {
            let child = walk_expr(inner, map, nodes);
            Some(push(
                nodes,
                "UnaryExpressionAst",
                expr.span,
                child.into_iter().collect(),
                vec![("TokenKind".to_string(), Member::Str(unary_token_name(text)))],
            ))
        }
        _ => None,
    }
}

fn assignment_operator(text: &str) -> String {
    for op in [
        "??=", "||=", "&&=", "<<=", ">>=", "**=", "+=", "-=", "*=", "/=", "%=", "=",
    ] {
        if text.contains(op) {
            return op.to_string();
        }
    }
    "=".to_string()
}

fn operator_name(op: &str) -> String {
    match op {
        "=" => "Equals",
        "+=" => "PlusEquals",
        "-=" => "MinusEquals",
        "*=" => "MultiplyEquals",
        "/=" => "DivideEquals",
        "%=" => "RemainderEquals",
        "??=" => "QuestionQuestionEquals",
        other => other,
    }
    .to_string()
}

fn binary_operator_name(text: &str) -> String {
    let lower = text.to_ascii_lowercase();
    for (needle, name) in [
        (" -isnot ", "IsNot"),
        (" -is ", "Is"),
        (" -as ", "As"),
        (" -eq ", "Ieq"),
        (" -ne ", "Ine"),
        (" -gt ", "Igt"),
        (" -ge ", "Ige"),
        (" -lt ", "Ilt"),
        (" -le ", "Ile"),
        (" + ", "Plus"),
        (" - ", "Minus"),
        (" * ", "Multiply"),
        (" / ", "Divide"),
        (" % ", "Rem"),
    ] {
        if lower.contains(needle) {
            return name.to_string();
        }
    }
    "Plus".to_string()
}

fn unary_token_name(text: &str) -> String {
    let lower = text.trim_start().to_ascii_lowercase();
    if lower.starts_with("-not") || lower.starts_with('!') {
        "Not".to_string()
    } else if lower.starts_with('-') {
        "Minus".to_string()
    } else {
        "Plus".to_string()
    }
}

fn string_constant_kind(text: &str, single: bool) -> String {
    let trimmed = text.trim_start();
    if trimmed.starts_with("@'") {
        "SingleQuotedHereString".to_string()
    } else if trimmed.starts_with("@\"") {
        "DoubleQuotedHereString".to_string()
    } else if single {
        "SingleQuoted".to_string()
    } else if trimmed.starts_with('"') {
        "DoubleQuoted".to_string()
    } else {
        "BareWord".to_string()
    }
}

fn foreach_variable(text: &str) -> Option<String> {
    let lower = text.to_ascii_lowercase();
    let hit = lower.find("foreach")?;
    let tail = &text[hit + "foreach".len()..];
    let start = tail.find('$')?;
    let rest = &tail[start + 1..];
    let name: String = rest
        .chars()
        .take_while(|c| c.is_ascii_alphanumeric() || *c == '_' || *c == ':')
        .collect();
    (!name.is_empty()).then_some(name)
}

fn command_name(callee: &Expression, text: &str) -> String {
    if let Some(spelled) = text.split_whitespace().next().filter(|s| !s.is_empty()) {
        return spelled.trim_start_matches('&').trim_start_matches('.').to_string();
    }
    if let ExprKind::Ident(name) = &callee.kind {
        return name.clone();
    }
    String::new()
}

fn literal_text(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(text)) => Some(text.clone()),
        ExprKind::Ident(name) => Some(name.clone()),
        _ => None,
    }
}

fn source_bracketed_type(text: &str) -> Option<String> {
    let trimmed = text.trim_start();
    if !trimmed.starts_with('[') {
        return None;
    }
    let end = trimmed.find(']')?;
    if trimmed.get(end + 1..)?.trim().is_empty() {
        Some(trimmed[1..end].trim().to_string())
    } else {
        None
    }
}

fn leading_bracketed_type(text: &str) -> Option<String> {
    let trimmed = text.trim_start();
    if !trimmed.starts_with('[') {
        return None;
    }
    let end = trimmed.find(']')?;
    let name = trimmed[1..end].trim();
    (!name.is_empty()).then(|| name.to_string())
}

fn source_static_member_ast(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> Option<usize> {
    let trimmed = text.trim_start();
    if !trimmed.starts_with('[') {
        return None;
    }
    let close = trimmed.find(']')?;
    let after = trimmed.get(close + 1..)?.trim_start();
    let member_text = after.strip_prefix("::")?;
    let type_name = trimmed[1..close].trim();
    let member_name: String = member_text
        .chars()
        .take_while(|c| c.is_ascii_alphanumeric() || *c == '_' || *c == '-')
        .collect();
    if member_name.is_empty() {
        return None;
    }
    let target = type_expression_node(type_name, span, nodes);
    let member = string_constant_node(&member_name, "BareWord", span, nodes);
    let ty = if member_text[member_name.len()..].trim_start().starts_with('(') {
        "InvokeMemberExpressionAst"
    } else {
        "MemberExpressionAst"
    };
    Some(push(
        nodes,
        ty,
        span,
        vec![target, member],
        vec![
            ("Expression".to_string(), Member::Node(target)),
            ("Member".to_string(), Member::Node(member)),
        ],
    ))
}

fn source_static_type_from_member_text(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> Option<usize> {
    let trimmed = text.trim_start();
    if !trimmed.starts_with('[') {
        return None;
    }
    let close = trimmed.find(']')?;
    trimmed.get(close + 1..)?.trim_start().starts_with("::").then(|| {
        type_expression_node(trimmed[1..close].trim(), span, nodes)
    })
}

fn source_member_ast(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> Option<usize> {
    let trimmed = text.trim();
    if trimmed.starts_with('[') || trimmed.contains("::") || trimmed.contains('|') {
        return None;
    }
    let dot = trimmed.rfind('.')?;
    if dot == 0 || trimmed[..dot].contains(char::is_whitespace) {
        return None;
    }
    let member_text = &trimmed[dot + 1..];
    let member_name: String = member_text
        .chars()
        .take_while(|c| c.is_ascii_alphanumeric() || *c == '_' || *c == '-')
        .collect();
    if member_name.is_empty() {
        return None;
    }
    let target_name = trimmed[..dot].trim().trim_start_matches('$');
    let target = push(
        nodes,
        "VariableExpressionAst",
        span,
        Vec::new(),
        vec![("VariablePath".to_string(), Member::Expr(variable_path_expr(target_name)))],
    );
    let member = string_constant_node(&member_name, "BareWord", span, nodes);
    Some(push(
        nodes,
        "MemberExpressionAst",
        span,
        vec![target, member],
        vec![
            ("Expression".to_string(), Member::Node(target)),
            ("Member".to_string(), Member::Node(member)),
        ],
    ))
}

fn type_expression_node(type_name: &str, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    push(
        nodes,
        "TypeExpressionAst",
        span,
        Vec::new(),
        vec![("TypeName".to_string(), Member::Expr(type_name_expr(type_name)))],
    )
}

fn type_constraint_node(type_name: &str, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    let type_expr = type_expression_node(type_name, span, nodes);
    push(
        nodes,
        "TypeConstraintAst",
        span,
        vec![type_expr],
        vec![
            ("TypeName".to_string(), Member::Expr(type_name_expr(type_name))),
            ("Type".to_string(), Member::Node(type_expr)),
        ],
    )
}

fn string_constant_node(value: &str, kind: &str, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    push(
        nodes,
        "StringConstantExpressionAst",
        span,
        Vec::new(),
        vec![
            ("Value".to_string(), Member::Str(value.to_string())),
            ("StringConstantType".to_string(), Member::Str(kind.to_string())),
            ("StaticType".to_string(), Member::Expr(type_ref_expr("System.String", false))),
        ],
    )
}

fn source_switch_keyword(text: &str) -> Option<usize> {
    let trimmed = text.trim_start();
    if trimmed.to_ascii_lowercase().starts_with("switch") {
        return Some(text.len() - trimmed.len());
    }
    if let Some(rest) = trimmed.strip_prefix(':') {
        let label_len = rest
            .chars()
            .take_while(|c| c.is_ascii_alphanumeric() || *c == '_' || *c == '-')
            .map(char::len_utf8)
            .sum::<usize>();
        let after = rest.get(label_len..)?.trim_start();
        if after.to_ascii_lowercase().starts_with("switch") {
            return text.find("switch").or_else(|| text.find("Switch"));
        }
    }
    None
}

fn source_compound_node(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> Option<usize> {
    if source_switch_keyword(text).is_some() {
        return Some(source_switch_node(text, span, nodes));
    }
    if source_loop_keyword(text, "foreach").is_some() {
        return Some(source_foreach_node(text, span, nodes));
    }
    if source_loop_keyword(text, "while").is_some() {
        return Some(source_while_node(text, span, nodes));
    }
    if source_loop_keyword(text, "for").is_some() {
        return Some(source_for_node(text, span, nodes));
    }
    if source_loop_keyword(text, "do").is_some() {
        return Some(source_do_node(text, span, nodes));
    }
    if starts_keyword(text.trim_start(), "try") {
        return Some(source_try_node(text, span, nodes));
    }
    if starts_keyword(text.trim_start(), "trap") {
        return Some(source_trap_node(text, span, nodes));
    }
    None
}

fn source_loop_keyword(text: &str, keyword: &str) -> Option<usize> {
    let trimmed = text.trim_start();
    if starts_keyword(trimmed, keyword) {
        return Some(text.len() - trimmed.len());
    }
    if let Some(rest) = trimmed.strip_prefix(':') {
        let label_len = rest
            .chars()
            .take_while(|c| c.is_ascii_alphanumeric() || *c == '_' || *c == '-')
            .map(char::len_utf8)
            .sum::<usize>();
        let after = rest.get(label_len..)?.trim_start();
        if starts_keyword(after, keyword) {
            return Some(text.len() - after.len());
        }
    }
    None
}

fn starts_keyword(text: &str, keyword: &str) -> bool {
    let lower = text.to_ascii_lowercase();
    if !lower.starts_with(keyword) {
        return false;
    }
    text.get(keyword.len()..)
        .and_then(|tail| tail.chars().next())
        .is_none_or(|ch| ch.is_whitespace() || ch == '(' || ch == '{')
}

fn source_loop_label(text: &str) -> String {
    source_switch_label(text)
}

fn source_switch_label(text: &str) -> String {
    let trimmed = text.trim_start();
    let Some(rest) = trimmed.strip_prefix(':') else {
        return String::new();
    };
    rest.chars()
        .take_while(|c| c.is_ascii_alphanumeric() || *c == '_' || *c == '-')
        .collect()
}

fn source_foreach_node(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    let header = source_parenthesized_after_keyword(text, "foreach").unwrap_or_default();
    let (var_name, cond_text) = split_foreach_header(&header);
    let variable = push(
        nodes,
        "VariableExpressionAst",
        span,
        Vec::new(),
        vec![
            ("VariablePath".to_string(), Member::Expr(variable_path_expr(&var_name))),
            ("__WholeText".to_string(), Member::Str(format!("${var_name}"))),
        ],
    );
    let condition = source_expression_node(&cond_text, span, nodes);
    let body = source_body_block_after_keyword(text, "foreach", span, nodes);
    push(
        nodes,
        "ForEachStatementAst",
        span,
        vec![variable, condition, body],
        vec![
            ("Variable".to_string(), Member::Node(variable)),
            ("Condition".to_string(), Member::Node(condition)),
            ("Body".to_string(), Member::Node(body)),
            ("Label".to_string(), Member::Str(source_loop_label(text))),
        ],
    )
}

fn source_while_node(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    let condition_text = source_parenthesized_after_keyword(text, "while").unwrap_or_default();
    let condition = source_expression_node(&condition_text, span, nodes);
    let body = source_body_block_after_keyword(text, "while", span, nodes);
    push(
        nodes,
        "WhileStatementAst",
        span,
        vec![condition, body],
        vec![
            ("Condition".to_string(), Member::Node(condition)),
            ("Body".to_string(), Member::Node(body)),
            ("Label".to_string(), Member::Str(source_loop_label(text))),
        ],
    )
}

fn source_for_node(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    let header = source_parenthesized_after_keyword(text, "for").unwrap_or_default();
    let parts = split_top_level_semicolon(&header);
    let initializer = parts.first().map(String::as_str).unwrap_or("").trim();
    let condition_text = parts.get(1).map(String::as_str).unwrap_or("").trim();
    let iterator = parts.get(2).map(String::as_str).unwrap_or("").trim();
    let init_node = (!initializer.is_empty()).then(|| source_expression_node(initializer, span, nodes));
    let cond_node = (!condition_text.is_empty()).then(|| source_expression_node(condition_text, span, nodes));
    let iter_node = (!iterator.is_empty()).then(|| source_expression_node(iterator, span, nodes));
    let body = source_body_block_after_keyword(text, "for", span, nodes);
    let mut children = Vec::new();
    children.extend(init_node);
    children.extend(cond_node);
    children.extend(iter_node);
    children.push(body);
    push(
        nodes,
        "ForStatementAst",
        span,
        children,
        vec![
            ("Initializer".to_string(), init_node.map(Member::Node).unwrap_or(Member::Null)),
            ("Condition".to_string(), cond_node.map(Member::Node).unwrap_or(Member::Null)),
            ("Iterator".to_string(), iter_node.map(Member::Node).unwrap_or(Member::Null)),
            ("Body".to_string(), Member::Node(body)),
            ("Label".to_string(), Member::Str(source_loop_label(text))),
        ],
    )
}

fn source_do_node(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    let lower = text.to_ascii_lowercase();
    let ty = if lower.contains(" until ") {
        "DoUntilStatementAst"
    } else {
        "DoWhileStatementAst"
    };
    let body = source_body_block_after_keyword(text, "do", span, nodes);
    let cond_key = if ty == "DoUntilStatementAst" { "until" } else { "while" };
    let condition_text = source_parenthesized_after_word(text, cond_key).unwrap_or_default();
    let condition = source_expression_node(&condition_text, span, nodes);
    push(
        nodes,
        ty,
        span,
        vec![body, condition],
        vec![
            ("Body".to_string(), Member::Node(body)),
            ("Condition".to_string(), Member::Node(condition)),
            ("Label".to_string(), Member::Str(source_loop_label(text))),
        ],
    )
}

fn source_try_node(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    let (try_body_text, mut cursor) = block_after_keyword_at(text, "try", 0)
        .map(|(_, body, end)| (body, end))
        .unwrap_or_else(|| (String::new(), 0));
    let body = source_statement_block_from_text(&try_body_text, span, nodes);
    let mut children = vec![body];
    let mut catches = Vec::new();
    let mut finally = None;
    loop {
        cursor = skip_ws(text, cursor);
        let rest = text.get(cursor..).unwrap_or("");
        if starts_keyword(rest, "catch") {
            if let Some((header, catch_body, end)) = block_after_keyword_at(text, "catch", cursor) {
                let type_names = bracketed_segments(&header);
                let mut catch_children = Vec::new();
                let catch_types = type_names
                    .into_iter()
                    .map(|ty| {
                        let node = type_constraint_node(&ty, span, nodes);
                        catch_children.push(node);
                        node
                    })
                    .collect::<Vec<_>>();
                let catch_body_node = source_statement_block_from_text(&catch_body, span, nodes);
                catch_children.push(catch_body_node);
                let catch_node = push(
                    nodes,
                    "CatchClauseAst",
                    span,
                    catch_children,
                    vec![
                        ("CatchTypes".to_string(), Member::Nodes(catch_types.clone())),
                        ("Body".to_string(), Member::Node(catch_body_node)),
                        ("IsCatchAll".to_string(), Member::Expr(Expression::bool(catch_types.is_empty()))),
                    ],
                );
                children.push(catch_node);
                catches.push(catch_node);
                cursor = end;
                continue;
            }
        }
        if starts_keyword(rest, "finally") {
            if let Some((_, finally_body, end)) = block_after_keyword_at(text, "finally", cursor) {
                let finally_node = source_statement_block_from_text(&finally_body, span, nodes);
                children.push(finally_node);
                finally = Some(finally_node);
                cursor = end;
                continue;
            }
        }
        break;
    }
    push(
        nodes,
        "TryStatementAst",
        span,
        children,
        vec![
            ("Body".to_string(), Member::Node(body)),
            ("CatchClauses".to_string(), Member::Nodes(catches)),
            ("Finally".to_string(), finally.map(Member::Node).unwrap_or(Member::Null)),
        ],
    )
}

fn source_trap_node(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    let header = text
        .trim_start()
        .strip_prefix("trap")
        .or_else(|| text.trim_start().strip_prefix("Trap"))
        .and_then(|rest| rest.split_once('{').map(|(head, _)| head.trim().to_string()))
        .unwrap_or_default();
    let trap_type = bracketed_segments(&header)
        .first()
        .map(|ty| type_constraint_node(ty, span, nodes));
    let body_text = text
        .find('{')
        .and_then(|open| find_matching_delim(text, open, '{', '}').map(|close| text[open + 1..close].to_string()))
        .unwrap_or_default();
    let body = source_statement_block_from_text(&body_text, span, nodes);
    let mut children = Vec::new();
    children.extend(trap_type);
    children.push(body);
    push(
        nodes,
        "TrapStatementAst",
        span,
        children,
        vec![
            ("TrapType".to_string(), trap_type.map(Member::Node).unwrap_or(Member::Null)),
            ("Body".to_string(), Member::Node(body)),
        ],
    )
}

fn block_after_keyword_at(text: &str, keyword: &str, start: usize) -> Option<(String, String, usize)> {
    let rest = text.get(start..)?;
    if !starts_keyword(rest.trim_start(), keyword) {
        return None;
    }
    let key_offset = start + rest.len() - rest.trim_start().len();
    let after_key = key_offset + keyword.len();
    let open = text.get(after_key..)?.find('{')? + after_key;
    let close = find_matching_delim(text, open, '{', '}')?;
    Some((
        text[after_key..open].trim().to_string(),
        text[open + 1..close].to_string(),
        close + 1,
    ))
}

fn split_foreach_header(header: &str) -> (String, String) {
    let lower = header.to_ascii_lowercase();
    if let Some(pos) = lower.find(" in ") {
        let var = header[..pos].trim().trim_start_matches('$').to_string();
        let cond = header[pos + 4..].trim().to_string();
        return (var, cond);
    }
    (String::new(), String::new())
}

fn source_parenthesized_after_keyword(text: &str, keyword: &str) -> Option<String> {
    let key = source_loop_keyword(text, keyword).or_else(|| source_switch_keyword(text))?;
    let tail = text.get(key + keyword.len()..)?;
    let open = tail.find('(')? + key + keyword.len();
    let close = find_matching_delim(text, open, '(', ')')?;
    Some(text[open + 1..close].trim().to_string())
}

fn source_body_block_after_keyword(text: &str, keyword: &str, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    let key = source_loop_keyword(text, keyword).or_else(|| source_switch_keyword(text)).unwrap_or(0);
    source_body_block_from(text, key, span, nodes)
}

fn source_body_block_from(text: &str, start: usize, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    if let Some(open) = text.get(start..).and_then(|tail| tail.find('{')).map(|pos| pos + start)
        && let Some(close) = find_matching_delim(text, open, '{', '}')
    {
        return source_statement_block_from_text(&text[open + 1..close], span, nodes);
    }
    source_statement_block_from_text("", span, nodes)
}

fn source_parenthesized_after_word(text: &str, keyword: &str) -> Option<String> {
    let lower = text.to_ascii_lowercase();
    let hit = lower.find(keyword)?;
    let tail = text.get(hit + keyword.len()..)?;
    let open = tail.find('(')? + hit + keyword.len();
    let close = find_matching_delim(text, open, '(', ')')?;
    Some(text[open + 1..close].trim().to_string())
}

fn split_top_level_semicolon(text: &str) -> Vec<String> {
    let mut out = Vec::new();
    let mut current = String::new();
    let mut quote: Option<char> = None;
    let mut square = 0i32;
    let mut paren = 0i32;
    for ch in text.chars() {
        if let Some(q) = quote {
            current.push(ch);
            if ch == q {
                quote = None;
            }
            continue;
        }
        match ch {
            '\'' | '"' => {
                quote = Some(ch);
                current.push(ch);
            }
            '[' => {
                square += 1;
                current.push(ch);
            }
            ']' => {
                square -= 1;
                current.push(ch);
            }
            '(' => {
                paren += 1;
                current.push(ch);
            }
            ')' => {
                paren -= 1;
                current.push(ch);
            }
            ';' if square == 0 && paren == 0 => {
                out.push(current.trim().to_string());
                current.clear();
            }
            _ => current.push(ch),
        }
    }
    out.push(current.trim().to_string());
    out
}

fn source_switch_node(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    let condition_text = source_switch_condition(text).unwrap_or_default();
    let condition = source_expression_node(&condition_text, span, nodes);
    let mut children = vec![condition];
    let mut clauses = Vec::new();
    let mut default = None;

    if let Some(body) = source_switch_body(text) {
        for (pattern, action) in split_switch_clauses(&body) {
            let action_block = source_statement_block_from_text(&action, span, nodes);
            children.push(action_block);
            if pattern.eq_ignore_ascii_case("default") {
                default = Some(action_block);
                continue;
            }
            let pattern_node = source_expression_node(&pattern, span, nodes);
            children.push(pattern_node);
            clauses.push(tuple2_expr(
                Expression::ident(&temp(pattern_node)),
                Expression::ident(&temp(action_block)),
            ));
        }
    }

    push(
        nodes,
        "SwitchStatementAst",
        span,
        children,
        vec![
            ("Condition".to_string(), Member::Node(condition)),
            ("Clauses".to_string(), Member::Expr(array_of(clauses))),
            ("Default".to_string(), default.map(Member::Node).unwrap_or(Member::Null)),
            ("Flags".to_string(), Member::Int(switch_flags(text))),
            ("Label".to_string(), Member::Str(source_switch_label(text))),
        ],
    )
}

fn source_switch_condition(text: &str) -> Option<String> {
    let switch_at = source_switch_keyword(text)?;
    let after_switch = text.get(switch_at + "switch".len()..)?;
    let open = after_switch.find('(')?;
    let start = switch_at + "switch".len() + open;
    let close = find_matching_delim(text, start, '(', ')')?;
    Some(text[start + 1..close].trim().to_string())
}

fn source_switch_body(text: &str) -> Option<String> {
    let switch_at = source_switch_keyword(text)?;
    let open = text.get(switch_at..)?.find('{')? + switch_at;
    let close = find_matching_delim(text, open, '{', '}')?;
    Some(text[open + 1..close].to_string())
}

fn find_matching_delim(text: &str, open: usize, left: char, right: char) -> Option<usize> {
    let mut depth = 0i32;
    let mut quote: Option<char> = None;
    for (idx, ch) in text.char_indices().filter(|(idx, _)| *idx >= open) {
        if let Some(q) = quote {
            if ch == q {
                quote = None;
            }
            continue;
        }
        match ch {
            '\'' | '"' => quote = Some(ch),
            c if c == left => depth += 1,
            c if c == right => {
                depth -= 1;
                if depth == 0 {
                    return Some(idx);
                }
            }
            _ => {}
        }
    }
    None
}

fn split_switch_clauses(body: &str) -> Vec<(String, String)> {
    let mut out = Vec::new();
    let mut i = 0usize;
    while i < body.len() {
        i = skip_ws(body, i);
        if i >= body.len() {
            break;
        }
        let pattern_start = i;
        if body[i..].to_ascii_lowercase().starts_with("default") {
            i += "default".len();
        } else if body[i..].starts_with('{') {
            let Some(end) = find_matching_delim(body, i, '{', '}') else {
                break;
            };
            i = end + 1;
        } else if body[i..].starts_with('\'') || body[i..].starts_with('"') {
            let quote = body[i..].chars().next().unwrap();
            i += quote.len_utf8();
            while i < body.len() {
                let ch = body[i..].chars().next().unwrap();
                i += ch.len_utf8();
                if ch == quote {
                    break;
                }
            }
        } else {
            while i < body.len() {
                let ch = body[i..].chars().next().unwrap();
                if ch.is_whitespace() || ch == '{' {
                    break;
                }
                i += ch.len_utf8();
            }
        }
        let pattern = body[pattern_start..i].trim().to_string();
        i = skip_ws(body, i);
        if i >= body.len() || !body[i..].starts_with('{') {
            break;
        }
        let Some(end) = find_matching_delim(body, i, '{', '}') else {
            break;
        };
        out.push((pattern, body[i + 1..end].to_string()));
        i = end + 1;
    }
    out
}

fn skip_ws(text: &str, mut i: usize) -> usize {
    while i < text.len() {
        let ch = text[i..].chars().next().unwrap();
        if !ch.is_whitespace() {
            break;
        }
        i += ch.len_utf8();
    }
    i
}

fn source_expression_node(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    let trimmed = text.trim();
    let member_text = |members: Vec<(String, Member)>| {
        let mut members = members;
        members.push(("__WholeText".to_string(), Member::Str(trimmed.to_string())));
        members
    };
    if trimmed.starts_with("$(") && trimmed.ends_with(')') {
        let child_text = trimmed[2..trimmed.len().saturating_sub(1)].trim();
        let child = source_expression_node(child_text, span, nodes);
        return push(
            nodes,
            "SubExpressionAst",
            span,
            vec![child],
            member_text(vec![("SubExpression".to_string(), Member::Node(child))]),
        );
    }
    if let Some((target_text, index_text)) = split_index_expression(trimmed) {
        let target = source_expression_node(target_text, span, nodes);
        let index = source_expression_node(index_text, span, nodes);
        return push(
            nodes,
            "IndexExpressionAst",
            span,
            vec![target, index],
            member_text(vec![
                ("Target".to_string(), Member::Node(target)),
                ("Index".to_string(), Member::Node(index)),
            ]),
        );
    }
    if let Some((target_text, member_name)) = split_member_expression(trimmed) {
        let target = source_expression_node(target_text, span, nodes);
        let member = string_constant_node(member_name, "BareWord", span, nodes);
        return push(
            nodes,
            "MemberExpressionAst",
            span,
            vec![target, member],
            member_text(vec![
                ("Expression".to_string(), Member::Node(target)),
                ("Member".to_string(), Member::Node(member)),
            ]),
        );
    }
    if trimmed.starts_with('{') {
        return push(
            nodes,
            "ScriptBlockExpressionAst",
            span,
            Vec::new(),
            member_text(Vec::new()),
        );
    }
    if trimmed.starts_with("@{") && trimmed.ends_with('}') {
        let inner = &trimmed[2..trimmed.len().saturating_sub(1)];
        let (children, pairs) = hashtable_child_nodes(inner, span, nodes);
        return push(
            nodes,
            "HashtableAst",
            span,
            children.clone(),
            member_text(vec![("KeyValuePairs".to_string(), Member::Expr(array_of(pairs)))]),
        );
    }
    if trimmed.starts_with("@(") && trimmed.ends_with(')') {
        let inner = &trimmed[2..trimmed.len().saturating_sub(1)];
        let elements = split_top_level_args(inner)
            .into_iter()
            .filter(|item| !item.trim().is_empty())
            .map(|item| source_expression_node(&item, span, nodes))
            .collect::<Vec<_>>();
        return push(
            nodes,
            "ArrayExpressionAst",
            span,
            elements.clone(),
            member_text(vec![("Elements".to_string(), Member::Nodes(elements))]),
        );
    }
    if top_level_comma_count(trimmed) > 0 {
        let elements = split_top_level_args(trimmed)
            .into_iter()
            .filter(|item| !item.trim().is_empty())
            .map(|item| source_expression_node(&item, span, nodes))
            .collect::<Vec<_>>();
        return push(
            nodes,
            "ArrayLiteralAst",
            span,
            elements.clone(),
            member_text(vec![("Elements".to_string(), Member::Nodes(elements))]),
        );
    }
    if (trimmed.starts_with('\'') && trimmed.ends_with('\'')) || (trimmed.starts_with('"') && trimmed.ends_with('"')) {
        let single = trimmed.starts_with('\'');
        return push(
            nodes,
            "StringConstantExpressionAst",
            span,
            Vec::new(),
            member_text(vec![
                ("Value".to_string(), Member::Str(unquote_ps_string(trimmed, single))),
                ("StringConstantType".to_string(), Member::Str(string_constant_kind(trimmed, single))),
                ("StaticType".to_string(), Member::Expr(type_ref_expr("System.String", false))),
            ]),
        );
    }
    if let Some((type_name, member_text_raw, is_invoke)) = split_static_member_expression(trimmed) {
        let target = type_expression_node(&type_name, span, nodes);
        let member = string_constant_node(&member_text_raw, "BareWord", span, nodes);
        return push(
            nodes,
            if is_invoke { "InvokeMemberExpressionAst" } else { "MemberExpressionAst" },
            span,
            vec![target, member],
            member_text(vec![
                ("Expression".to_string(), Member::Node(target)),
                ("Member".to_string(), Member::Node(member)),
                ("Arguments".to_string(), Member::Nodes(Vec::new())),
            ]),
        );
    }
    if trimmed.starts_with('[') && trimmed.ends_with(']') {
        return type_expression_node(&trimmed[1..trimmed.len().saturating_sub(1)], span, nodes);
    }
    if let Ok(value) = trimmed.parse::<i64>() {
        return push(
            nodes,
            "ConstantExpressionAst",
            span,
            Vec::new(),
            member_text(vec![("Value".to_string(), Member::Int(value))]),
        );
    }
    if trimmed.contains(" + ")
        || trimmed.contains(" - ")
        || trimmed.contains(" * ")
        || trimmed.contains(" / ")
        || trimmed.contains(" % ")
        || trimmed.contains(" -gt ")
        || trimmed.contains(" -is ")
        || trimmed.contains(" -as ")
    {
        let (left, right) = binary_operands(trimmed);
        let left_node = left.map(|value| source_expression_node(value, span, nodes));
        let right_node = right.map(|value| source_expression_node(value, span, nodes));
        let mut children = Vec::new();
        children.extend(left_node);
        children.extend(right_node);
        return push(
            nodes,
            "BinaryExpressionAst",
            span,
            children,
            member_text(vec![
                ("Left".to_string(), left_node.map(Member::Node).unwrap_or(Member::Null)),
                ("Right".to_string(), right_node.map(Member::Node).unwrap_or(Member::Null)),
                ("Operator".to_string(), Member::Str(binary_operator_name(trimmed))),
            ]),
        );
    }
    push(
        nodes,
        "VariableExpressionAst",
        span,
        Vec::new(),
        member_text(vec![(
            "VariablePath".to_string(),
            Member::Expr(variable_path_expr(trimmed.trim_start_matches('$'))),
        )]),
    )
}

fn unquote_ps_string(text: &str, single: bool) -> String {
    let inner = &text[1..text.len().saturating_sub(1)];
    if single {
        return inner.replace("''", "'");
    }
    let mut out = String::new();
    let mut chars = inner.chars();
    while let Some(ch) = chars.next() {
        if ch != '`' {
            out.push(ch);
            continue;
        }
        match chars.next() {
            Some('0') => out.push('\0'),
            Some('a') => out.push('\x07'),
            Some('b') => out.push('\x08'),
            Some('f') => out.push('\x0c'),
            Some('n') => out.push('\n'),
            Some('r') => out.push('\r'),
            Some('t') => out.push('\t'),
            Some('v') => out.push('\x0b'),
            Some(next) => out.push(next),
            None => out.push('`'),
        }
    }
    out
}

fn hashtable_child_nodes(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> (Vec<usize>, Vec<Expression>) {
    let mut children = Vec::new();
    let mut pairs = Vec::new();
    for entry in split_action_statements(text) {
        let Some((key, value)) = split_top_level_equals(&entry) else {
            continue;
        };
        let key_node = source_expression_node(key.trim(), span, nodes);
        let value_node = source_expression_node(value.trim(), span, nodes);
        children.push(key_node);
        children.push(value_node);
        pairs.push(tuple2_expr(Expression::ident(&temp(key_node)), Expression::ident(&temp(value_node))));
    }
    (children, pairs)
}

fn top_level_comma_count(text: &str) -> usize {
    let mut quote: Option<char> = None;
    let mut square = 0i32;
    let mut paren = 0i32;
    let mut brace = 0i32;
    let mut count = 0usize;
    for ch in text.chars() {
        if let Some(q) = quote {
            if ch == q {
                quote = None;
            }
            continue;
        }
        match ch {
            '\'' | '"' => quote = Some(ch),
            '[' => square += 1,
            ']' => square -= 1,
            '(' => paren += 1,
            ')' => paren -= 1,
            '{' => brace += 1,
            '}' => brace -= 1,
            ',' if square == 0 && paren == 0 && brace == 0 => count += 1,
            _ => {}
        }
    }
    count
}

fn split_static_member_expression(text: &str) -> Option<(String, String, bool)> {
    let rest = text.strip_prefix('[')?;
    let close = rest.find(']')?;
    let type_name = rest[..close].trim();
    let after_type = rest[close + 1..].trim_start();
    let after_scope = after_type.strip_prefix("::")?.trim_start();
    let member: String = after_scope
        .chars()
        .take_while(|ch| ch.is_ascii_alphanumeric() || *ch == '_')
        .collect();
    if type_name.is_empty() || member.is_empty() {
        return None;
    }
    let is_invoke = after_scope[member.len()..].trim_start().starts_with('(');
    Some((type_name.to_string(), member, is_invoke))
}

fn split_index_expression(text: &str) -> Option<(&str, &str)> {
    if text.starts_with('[') {
        return None;
    }
    let open = text.rfind('[')?;
    if !text.ends_with(']') {
        return None;
    }
    let close = find_matching_delim(text, open, '[', ']')?;
    if close != text.len() - 1 || text[..open].trim().is_empty() {
        return None;
    }
    Some((text[..open].trim(), text[open + 1..close].trim()))
}

fn split_member_expression(text: &str) -> Option<(&str, &str)> {
    if text.starts_with('[') {
        return None;
    }
    let dot = text.rfind('.')?;
    let target = text[..dot].trim();
    let member = text[dot + 1..].trim();
    if target.is_empty()
        || member.is_empty()
        || !target.starts_with('$')
        || !member.chars().all(|ch| ch.is_ascii_alphanumeric() || ch == '_')
    {
        return None;
    }
    Some((target, member))
}

fn binary_operands(text: &str) -> (Option<&str>, Option<&str>) {
    for op in [" -is ", " -as ", " -gt ", " + ", " - "] {
        if let Some(index) = text.find(op) {
            return (Some(text[..index].trim()), Some(text[index + op.len()..].trim()));
        }
    }
    (None, None)
}

fn source_statement_block_from_text(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    let statements: Vec<usize> = split_action_statements(text)
        .into_iter()
        .filter_map(|stmt| source_statement_node(&stmt, span, nodes))
        .collect();
    push(
        nodes,
        "StatementBlockAst",
        span,
        statements.clone(),
        vec![
            ("Statements".to_string(), Member::Nodes(statements)),
            ("__WholeText".to_string(), Member::Str(text.trim().to_string())),
        ],
    )
}

fn source_statement_node(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> Option<usize> {
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return None;
    }
    if let Some(index) = source_compound_node(trimmed, span, nodes) {
        return Some(index);
    }
    if trimmed.to_ascii_lowercase().starts_with("if ") || trimmed.to_ascii_lowercase().starts_with("if(") {
        let body = source_body_block_from(trimmed, 0, span, nodes);
        return Some(push(
            nodes,
            "IfStatementAst",
            span,
            vec![body],
            vec![
                ("Clauses".to_string(), Member::Expr(array_of(vec![tuple2_expr(Expression::null(), Expression::ident(&temp(body)))]))),
                ("ElseClause".to_string(), Member::Null),
            ],
        ));
    }
    if trimmed.eq_ignore_ascii_case("break") || trimmed.to_ascii_lowercase().starts_with("break ") {
        return Some(flow_statement_node("BreakStatementAst", trimmed, span, nodes));
    }
    if trimmed.eq_ignore_ascii_case("continue") || trimmed.to_ascii_lowercase().starts_with("continue ") {
        return Some(flow_statement_node("ContinueStatementAst", trimmed, span, nodes));
    }
    if trimmed.eq_ignore_ascii_case("return") || trimmed.to_ascii_lowercase().starts_with("return ") {
        return Some(push(nodes, "ReturnStatementAst", span, Vec::new(), Vec::new()));
    }
    if trimmed.eq_ignore_ascii_case("throw") || trimmed.to_ascii_lowercase().starts_with("throw ") {
        return Some(source_throw_node(trimmed, span, nodes));
    }
    if trimmed.contains('|') {
        return Some(pipeline_ast_node(trimmed, span, nodes));
    }
    if let Some(assign) = source_assignment_node(trimmed, span, nodes) {
        return Some(assign);
    }
    if source_looks_like_command(trimmed) {
        return Some(command_ast_node(trimmed, span, nodes));
    }
    Some(source_expression_node(trimmed, span, nodes))
}

fn source_throw_node(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    let payload = text
        .trim_start()
        .strip_prefix("throw")
        .or_else(|| text.trim_start().strip_prefix("Throw"))
        .map(str::trim)
        .unwrap_or("");
    let pipeline = if payload.is_empty() {
        None
    } else {
        Some(pipeline_ast_node(payload, span, nodes))
    };
    push(
        nodes,
        "ThrowStatementAst",
        span,
        pipeline.into_iter().collect(),
        vec![("Pipeline".to_string(), pipeline.map(Member::Node).unwrap_or(Member::Null))],
    )
}

fn source_assignment_node(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> Option<usize> {
    let (left_text, operator, right_text) = split_top_level_assignment(text)?;
    let left_text = left_text.trim();
    let left_leaf = if let Some(close) = left_text.trim_start().strip_prefix('[').and_then(|rest| rest.find(']').map(|idx| idx + 1)) {
        source_expression_node(left_text.trim_start()[close..].trim(), span, nodes)
    } else {
        source_expression_node(left_text, span, nodes)
    };
    let left = typed_assignment_left(left_text, Some(left_leaf), span, nodes).unwrap_or(left_leaf);
    let right = assignment_right_node(right_text.trim(), span, nodes);
    let error_position = assignment_error_position(operator);
    Some(push(
        nodes,
        "AssignmentStatementAst",
        span,
        vec![left, right],
        vec![
            ("Left".to_string(), Member::Node(left)),
            ("Right".to_string(), Member::Node(right)),
            ("Operator".to_string(), Member::Str(assignment_operator_name(operator))),
            ("ErrorPosition".to_string(), Member::Expr(error_position)),
        ],
    ))
}

fn assignment_right_node(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    if let Some(assign) = source_assignment_node(text, span, nodes) {
        return assign;
    }
    if text.contains('|') || source_looks_like_command(text) {
        pipeline_ast_node(text, span, nodes)
    } else {
        command_expression_node(source_expression_node(text, span, nodes), Vec::new(), span, nodes)
    }
}

fn split_top_level_assignment(text: &str) -> Option<(&str, &str, &str)> {
    let mut quote: Option<char> = None;
    let mut square = 0i32;
    let mut paren = 0i32;
    let mut brace = 0i32;
    for (idx, ch) in text.char_indices() {
        if let Some(q) = quote {
            if ch == q {
                quote = None;
            }
            continue;
        }
        match ch {
            '\'' | '"' => quote = Some(ch),
            '[' => square += 1,
            ']' => square -= 1,
            '(' => paren += 1,
            ')' => paren -= 1,
            '{' => brace += 1,
            '}' => brace -= 1,
            '=' if square == 0 && paren == 0 && brace == 0 => {
                let prefix = &text[..idx + 1];
                let op_start = ["??=", "||=", "&&=", "<<=", ">>=", "**=", "+=", "-=", "*=", "/=", "%=", "="]
                    .into_iter()
                    .find_map(|op| prefix.ends_with(op).then_some(idx + 1 - op.len()))
                    .unwrap_or(idx);
                let left = &text[..op_start];
                let op = &text[op_start..idx + 1];
                let right = &text[idx + 1..];
                if !left.trim().is_empty() {
                    return Some((left, op, right));
                }
            }
            _ => {}
        }
    }
    None
}

fn assignment_operator_name(operator: &str) -> String {
    match operator.trim() {
        "??=" => "QuestionQuestionEquals",
        "||=" => "BarBarEquals",
        "&&=" => "AmpersandAmpersandEquals",
        "<<=" => "ShiftLeftEquals",
        ">>=" => "ShiftRightEquals",
        "**=" => "PowerEquals",
        "+=" => "PlusEquals",
        "-=" => "MinusEquals",
        "*=" => "MultiplyEquals",
        "/=" => "DivideEquals",
        "%=" => "RemainderEquals",
        _ => "Equals",
    }
    .to_string()
}

fn assignment_error_position(operator: &str) -> Expression {
    Expression::new(ExprKind::Object(vec![
        prop("Text", Expression::string(operator.trim())),
        prop("StartOffset", Expression::int(0)),
        prop("EndOffset", Expression::int(operator.trim().len() as i64)),
    ]))
}

fn flow_statement_node(ty: &str, text: &str, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    let label_text = text.split_whitespace().nth(1).unwrap_or("").trim_start_matches(':');
    let label = if label_text.is_empty() {
        None
    } else {
        Some(string_constant_node(label_text, "BareWord", span, nodes))
    };
    push(
        nodes,
        ty,
        span,
        label.into_iter().collect(),
        vec![("Label".to_string(), label.map(Member::Node).unwrap_or(Member::Null))],
    )
}

fn split_action_statements(text: &str) -> Vec<String> {
    let mut out = Vec::new();
    let mut current = String::new();
    let mut brace = 0i32;
    let mut quote: Option<char> = None;
    for ch in text.chars() {
        if let Some(q) = quote {
            current.push(ch);
            if ch == q {
                quote = None;
            }
            continue;
        }
        match ch {
            '\'' | '"' => {
                quote = Some(ch);
                current.push(ch);
            }
            '{' => {
                brace += 1;
                current.push(ch);
            }
            '}' => {
                brace -= 1;
                current.push(ch);
            }
            ';' | '\n' if brace == 0 => {
                if !current.trim().is_empty() {
                    out.push(current.trim().to_string());
                    current.clear();
                }
            }
            _ => current.push(ch),
        }
    }
    if !current.trim().is_empty() {
        out.push(current.trim().to_string());
    }
    out
}

fn switch_flags(text: &str) -> i64 {
    let mut flags = 0i64;
    for (word, bit) in [
        ("Regex", 1i64),
        ("Wildcard", 2),
        ("Exact", 4),
        ("CaseSensitive", 8),
        ("File", 16),
        ("Parallel", 32),
    ] {
        if text
            .to_lowercase()
            .contains(&format!("-{}", word.to_lowercase()))
        {
            flags |= bit;
        }
    }
    flags
}

// ── emission ──────────────────────────────────────────────────────────────

fn temp(index: usize) -> String {
    format!("__ps_ast_{index}")
}

fn prop(name: &str, value: Expression) -> ObjectProperty {
    ObjectProperty::KeyValue {
        key: Expression::string(name),
        value,
    }
}

fn array_of(items: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Array(
        items
            .into_iter()
            .map(|value| ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false,
            })
            .collect(),
    ))
}

fn tuple2_expr(first: Expression, second: Expression) -> Expression {
    Expression::new(ExprKind::Object(vec![
        prop("Item1", first),
        prop("Item2", second),
    ]))
}

fn type_ref_expr(full_name: &str, is_enum: bool) -> Expression {
    Expression::new(ExprKind::Object(vec![
        prop("FullName", Expression::string(full_name)),
        prop("Name", Expression::string(full_name.rsplit('.').next().unwrap_or(full_name))),
        prop("IsEnum", Expression::bool(is_enum)),
    ]))
}

fn variable_path_expr(path: &str) -> Expression {
    let (scope, user_path) = path
        .split_once(':')
        .map(|(s, p)| (s.to_ascii_lowercase(), p.to_string()))
        .unwrap_or_else(|| ("".to_string(), path.to_string()));
    let is_drive = !scope.is_empty()
        && !matches!(
            scope.as_str(),
            "global" | "local" | "private" | "script" | "variable"
        );
    Expression::new(ExprKind::Object(vec![
        prop("UserPath", Expression::string(&user_path)),
        prop("UnqualifiedPath", Expression::string(&user_path)),
        prop("DriveName", if is_drive { Expression::string(&scope) } else { Expression::null() }),
        prop("IsDriveQualified", Expression::bool(is_drive)),
        prop("IsGlobal", Expression::bool(scope == "global")),
        prop("IsLocal", Expression::bool(scope == "local")),
        prop("IsPrivate", Expression::bool(scope == "private")),
        prop("IsScript", Expression::bool(scope == "script")),
        prop("IsVariable", Expression::bool(scope == "variable")),
        prop("IsUnqualified", Expression::bool(scope.is_empty())),
    ]))
}

fn type_name_expr(name: &str) -> Expression {
    let (base, array_rank) = peel_array_type(name.trim());
    let generic_args = generic_args(&base);
    let generic = !generic_args.is_empty();
    let reflection_name = type_accelerator(&base).unwrap_or(base.as_str()).to_string();
    let element_type = if array_rank > 0 {
        let element = peel_array_type(&base).0;
        type_name_expr(&element)
    } else {
        Expression::null()
    };
    let reflection = type_ref_expr(&reflection_name, reflection_name.rsplit('.').next().is_some_and(|n| n.ends_with("Access")));
    Expression::new(ExprKind::Object(vec![
        prop("Name", Expression::string(base.rsplit('.').next().unwrap_or(&base))),
        prop("FullName", Expression::string(&base)),
        prop("IsArray", Expression::bool(array_rank > 0)),
        prop("Rank", Expression::int(array_rank.max(1) as i64)),
        prop("ElementType", element_type),
        prop("IsGeneric", Expression::bool(generic)),
        prop("GenericArguments", array_of(generic_args.into_iter().map(|arg| type_name_expr(&arg)).collect())),
        prop("GetReflectionType", lambda_expr(
            vec!["__self"],
            vec![Statement::new(StmtKind::Return(Some(reflection)))],
        )),
    ]))
}

fn peel_array_type(name: &str) -> (String, usize) {
    let mut base = name.trim().to_string();
    let mut rank = 0usize;
    loop {
        let Some(open) = base.rfind('[') else {
            break;
        };
        if !base.ends_with(']') {
            break;
        }
        let suffix = &base[open + 1..base.len() - 1];
        if suffix.chars().all(|c| c == ',') {
            rank = suffix.len() + 1;
            base.truncate(open);
            continue;
        }
        break;
    }
    (base, rank)
}

fn generic_args(name: &str) -> Vec<String> {
    let Some(open) = name.find('[') else {
        return Vec::new();
    };
    if !name.ends_with(']') {
        return Vec::new();
    }
    split_top_level_args(&name[open + 1..name.len() - 1])
}

fn type_accelerator(name: &str) -> Option<&'static str> {
    match name.to_ascii_lowercase().as_str() {
        "hashtable" => Some("System.Collections.Hashtable"),
        "pscustomobject" => Some("System.Management.Automation.PSObject"),
        "scriptblock" => Some("System.Management.Automation.ScriptBlock"),
        "switch" => Some("System.Management.Automation.SwitchParameter"),
        "string" => Some("System.String"),
        "int" => Some("System.Int32"),
        "long" => Some("System.Int64"),
        "double" => Some("System.Double"),
        "decimal" => Some("System.Decimal"),
        "bool" => Some("System.Boolean"),
        "datetime" => Some("System.DateTime"),
        _ => None,
    }
}

struct ParamInfo {
    name: String,
    attrs: Vec<String>,
    default: Option<String>,
    static_type: Option<String>,
}

fn param_block_node(source: &str, span: Span, nodes: &mut Vec<AstNode>) -> Option<usize> {
    let (attrs, params) = parse_param_block(source)?;
    let attr_nodes: Vec<usize> = attrs
        .into_iter()
        .map(|attr| attribute_node(&attr, span, nodes))
        .collect();
    let mut children = attr_nodes.clone();
    let mut param_nodes = Vec::new();
    for param_info in params {
        let mut param_attrs = Vec::new();
        let mut type_name = param_info.static_type.clone();
        for attr in param_info.attrs {
            let node = if is_type_constraint_name(&attr) {
                let clean = attr.split('(').next().unwrap_or(&attr).trim().to_string();
                type_name = Some(clean.clone());
                type_constraint_node(&clean, span, nodes)
            } else {
                attribute_node(&attr, span, nodes)
            };
            children.push(node);
            param_attrs.push(node);
        }
        let default = param_info
            .default
            .as_deref()
            .map(|value| source_expression_node(value, span, nodes));
        children.extend(default);
        let param = parameter_node(&param_info.name, span, param_attrs, default, type_name, nodes);
        children.push(param);
        param_nodes.push(param);
    }
    Some(push(
        nodes,
        "ParamBlockAst",
        span,
        children,
        vec![
            ("Attributes".to_string(), Member::Nodes(attr_nodes)),
            ("Parameters".to_string(), Member::Nodes(param_nodes)),
        ],
    ))
}

fn parse_param_block(source: &str) -> Option<(Vec<String>, Vec<ParamInfo>)> {
    let lower = source.to_ascii_lowercase();
    let start = lower.find("param(")?;
    let attrs = parse_block_attrs(&source[..start]);
    let body = &source[start + "param(".len()..];
    let end = find_matching_paren_body_end(body)?;
    let params = parse_param_names(&body[..end]);
    Some((attrs, params))
}

fn parse_block_attrs(prefix: &str) -> Vec<String> {
    let mut attrs = Vec::new();
    let mut rest = prefix.trim();
    while let Some(open) = rest.find('[') {
        let after = &rest[open + 1..];
        let Some(close) = after.find(']') else {
            break;
        };
        let raw = after[..close].trim();
        if !raw.is_empty() {
            attrs.push(raw.to_string());
        }
        rest = &after[close + 1..];
    }
    attrs
}

fn attribute_node(raw: &str, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    let name = raw.split(['(', ' ']).next().unwrap_or(raw).trim();
    let args = raw
        .find('(')
        .and_then(|open| raw.rfind(')').map(|close| raw[open + 1..close].to_string()));
    push(
        nodes,
        "AttributeAst",
        span,
        Vec::new(),
        vec![
            ("TypeName".to_string(), Member::Expr(type_name_expr(name))),
            ("PositionalArguments".to_string(), Member::Expr(array_of(attribute_positional_args(args.as_deref())))),
            ("NamedArguments".to_string(), Member::Expr(array_of(attribute_named_args(args.as_deref())))),
        ],
    )
}

fn attribute_positional_args(args: Option<&str>) -> Vec<Expression> {
    args.unwrap_or("")
        .split(',')
        .map(str::trim)
        .filter(|arg| !arg.is_empty() && !arg.contains('=') && !arg.chars().all(|c| c.is_ascii_alphabetic()))
        .map(attribute_argument_expr)
        .collect()
}

fn attribute_named_args(args: Option<&str>) -> Vec<Expression> {
    args.unwrap_or("")
        .split(',')
        .map(str::trim)
        .filter(|arg| !arg.is_empty())
        .filter_map(|arg| {
            if let Some((name, value)) = arg.split_once('=') {
                Some(named_attribute_argument_expr_full(name.trim(), value.trim(), false))
            } else if arg.chars().all(|c| c.is_ascii_alphabetic()) {
                Some(named_attribute_argument_expr_full(arg, "$true", true))
            } else {
                None
            }
        })
        .collect()
}

fn named_attribute_argument_expr_full(name: &str, value: &str, omitted: bool) -> Expression {
    Expression::new(ExprKind::Object(vec![
        prop("__type", Expression::string("NamedAttributeArgumentAst")),
        prop("ArgumentName", Expression::string(name)),
        prop("ExpressionOmitted", Expression::bool(omitted)),
        prop("Argument", source_expression_object(value)),
    ]))
}

fn source_expression_object(text: &str) -> Expression {
    let trimmed = text.trim();
    if trimmed.eq_ignore_ascii_case("$true") || trimmed.eq_ignore_ascii_case("true") {
        return Expression::new(ExprKind::Object(vec![
            prop("__type", Expression::string("VariableExpressionAst")),
            prop("VariablePath", variable_path_expr("true")),
        ]));
    }
    if trimmed.eq_ignore_ascii_case("$false") || trimmed.eq_ignore_ascii_case("false") {
        return Expression::new(ExprKind::Object(vec![
            prop("__type", Expression::string("VariableExpressionAst")),
            prop("VariablePath", variable_path_expr("false")),
        ]));
    }
    if let Ok(value) = trimmed.parse::<i64>() {
        return Expression::new(ExprKind::Object(vec![
            prop("__type", Expression::string("ConstantExpressionAst")),
            prop("Value", Expression::int(value)),
        ]));
    }
    Expression::new(ExprKind::Object(vec![
        prop("__type", Expression::string("StringConstantExpressionAst")),
        prop("Value", Expression::string(trimmed.trim_matches('"').trim_matches('\''))),
    ]))
}

fn parse_param_names(body: &str) -> Vec<ParamInfo> {
    if body.trim().is_empty() {
        return Vec::new();
    }
    let mut out = Vec::new();
    for part in split_top_level_args(body) {
        let attrs = bracketed_segments(&part);
        if let Some(dollar) = part.rfind('$') {
            let name: String = part[dollar + 1..]
                .chars()
                .take_while(|c| c.is_ascii_alphanumeric() || *c == '_')
                .collect();
            if !name.is_empty() {
                let default = split_top_level_equals(&part).map(|(_, value)| value.trim().to_string());
                let static_type = attrs
                    .iter()
                    .find(|attr| is_type_constraint_name(attr))
                    .map(|attr| attr.split('(').next().unwrap_or(attr).trim().to_string());
                out.push(ParamInfo {
                    name,
                    attrs,
                    default,
                    static_type,
                });
            }
        }
    }
    out
}

fn bracketed_segments(text: &str) -> Vec<String> {
    let mut out = Vec::new();
    let mut i = 0usize;
    while i < text.len() {
        let Some(open_rel) = text[i..].find('[') else {
            break;
        };
        let open = i + open_rel;
        let Some(close_rel) = text[open + 1..].find(']') else {
            break;
        };
        let close = open + 1 + close_rel;
        let raw = text[open + 1..close].trim();
        if !raw.is_empty() {
            out.push(raw.to_string());
        }
        i = close + 1;
    }
    out
}

fn is_type_constraint_name(raw: &str) -> bool {
    let name = raw.split('(').next().unwrap_or(raw).trim();
    let lower = name.to_ascii_lowercase();
    !lower.starts_with("parameter")
        && !lower.starts_with("alias")
        && !lower.starts_with("validate")
        && !lower.starts_with("cmdletbinding")
        && !lower.starts_with("outputtype")
}

fn split_top_level_equals(text: &str) -> Option<(&str, &str)> {
    let mut quote: Option<char> = None;
    let mut square = 0i32;
    let mut paren = 0i32;
    for (idx, ch) in text.char_indices() {
        if let Some(q) = quote {
            if ch == q {
                quote = None;
            }
            continue;
        }
        match ch {
            '\'' | '"' => quote = Some(ch),
            '[' => square += 1,
            ']' => square -= 1,
            '(' => paren += 1,
            ')' => paren -= 1,
            '=' if square == 0 && paren == 0 => return Some((&text[..idx], &text[idx + 1..])),
            _ => {}
        }
    }
    None
}

fn find_matching_paren_body_end(body: &str) -> Option<usize> {
    let mut depth = 1i32;
    let mut quote: Option<char> = None;
    for (idx, ch) in body.char_indices() {
        if let Some(q) = quote {
            if ch == q {
                quote = None;
            }
            continue;
        }
        match ch {
            '\'' | '"' => quote = Some(ch),
            '(' => depth += 1,
            ')' => {
                depth -= 1;
                if depth == 0 {
                    return Some(idx);
                }
            }
            _ => {}
        }
    }
    None
}

fn typed_assignment_left(
    text: &str,
    left: Option<usize>,
    span: Span,
    nodes: &mut Vec<AstNode>,
) -> Option<usize> {
    let trimmed = text.trim_start();
    if !trimmed.starts_with('[') {
        return left;
    }
    let end = trimmed.find(']')?;
    let type_name = &trimmed[1..end];
    let type_node = push(
        nodes,
        "TypeExpressionAst",
        span,
        Vec::new(),
        vec![("TypeName".to_string(), Member::Expr(type_name_expr(type_name)))],
    );
    let mut children = vec![type_node];
    children.extend(left);
    Some(push(
        nodes,
        "ConvertExpressionAst",
        span,
        children,
        vec![
            ("Type".to_string(), Member::Node(type_node)),
            ("Child".to_string(), left.map(Member::Node).unwrap_or(Member::Null)),
        ],
    ))
}

fn pipeline_element_nodes(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> Vec<usize> {
    text.split('|')
        .map(str::trim)
        .filter(|part| !part.is_empty())
        .map(|part| {
            if source_looks_like_command(part) || command_invocation_operator(part) != "Unknown" || has_redirection(part) {
                command_ast_node(part, span, nodes)
            } else {
                command_expression_node(source_expression_node(part, span, nodes), Vec::new(), span, nodes)
            }
        })
        .collect()
}

fn pipeline_ast_node(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    let elements = pipeline_element_nodes(text, span, nodes);
    let pure_expr = (elements.len() == 1)
        .then_some(elements[0])
        .and_then(|element| nodes.get(element))
        .filter(|node| node.ty == "CommandExpressionAst")
        .and_then(|node| node.members.iter().find_map(|(name, member)| {
            (name == "Expression").then_some(member).and_then(|member| match member {
                Member::Node(index) => Some(Expression::ident(&temp(*index))),
                _ => None,
            })
        }))
        .unwrap_or_else(Expression::null);
    push(
        nodes,
        "PipelineAst",
        span,
        elements.clone(),
        vec![
            ("PipelineElements".to_string(), Member::Nodes(elements)),
            ("GetPureExpression".to_string(), Member::Expr(function_expr(
                Vec::new(),
                vec![Statement::new(StmtKind::Return(Some(pure_expr)))],
            ))),
        ],
    )
}

fn command_expression_node(expr: usize, redirections: Vec<usize>, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    let mut children = vec![expr];
    children.extend(redirections.iter().copied());
    push(
        nodes,
        "CommandExpressionAst",
        span,
        children,
        vec![
            ("Expression".to_string(), Member::Node(expr)),
            ("Redirections".to_string(), Member::Nodes(redirections)),
        ],
    )
}

fn command_ast_node(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> usize {
    let invocation = command_invocation_operator(text);
    let command_text = strip_invocation_operator(text);
    let (clean_text, redirections) = command_redirection_nodes(command_text, span, nodes);
    let elements = command_element_nodes(&clean_text, span, nodes);
    let command_name = if invocation == "Unknown" {
        clean_text.split_whitespace().next().unwrap_or("").to_string()
    } else if clean_text.trim_start().starts_with('$') {
        String::new()
    } else {
        clean_text.split_whitespace().next().unwrap_or("").to_string()
    };
    let mut children = elements.clone();
    children.extend(redirections.iter().copied());
    push(
        nodes,
        "CommandAst",
        span,
        children,
        vec![
            ("CommandElements".to_string(), Member::Nodes(elements)),
            ("CommandName".to_string(), if command_name.is_empty() { Member::Null } else { Member::Str(command_name) }),
            ("InvocationOperator".to_string(), Member::Str(invocation.to_string())),
            ("Redirections".to_string(), Member::Nodes(redirections)),
        ],
    )
}

fn command_invocation_operator(text: &str) -> &'static str {
    let trimmed = text.trim_start();
    if trimmed.starts_with('&') {
        "Ampersand"
    } else if trimmed.starts_with('.') {
        "Dot"
    } else {
        "Unknown"
    }
}

fn strip_invocation_operator(text: &str) -> &str {
    let trimmed = text.trim_start();
    if trimmed.starts_with('&') || trimmed.starts_with('.') {
        trimmed[1..].trim_start()
    } else {
        text
    }
}

fn has_redirection(text: &str) -> bool {
    split_command_like_tokens(text)
        .into_iter()
        .any(|token| redirection_operator(&token).is_some() || merging_redirection(&token).is_some())
}

fn command_redirection_nodes(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> (String, Vec<usize>) {
    let tokens = split_command_like_tokens(text);
    let mut clean = Vec::new();
    let mut redirections = Vec::new();
    let mut i = 0usize;
    while i < tokens.len() {
        let token = tokens[i].as_str();
        if let Some((from, append, attached)) = redirection_operator(token) {
            let location_text = attached
                .filter(|value| !value.is_empty())
                .map(str::to_string)
                .or_else(|| {
                    let next = tokens.get(i + 1).cloned();
                    if next.is_some() {
                        i += 1;
                    }
                    next
                })
                .unwrap_or_default();
            let location = source_expression_node(&location_text, span, nodes);
            redirections.push(push(
                nodes,
                "FileRedirectionAst",
                span,
                vec![location],
                vec![
                    ("FromStream".to_string(), Member::Str(from.to_string())),
                    ("Append".to_string(), Member::Bool(append)),
                    ("Location".to_string(), Member::Node(location)),
                ],
            ));
        } else if let Some(from) = merging_redirection(token) {
            redirections.push(push(
                nodes,
                "MergingRedirectionAst",
                span,
                Vec::new(),
                vec![
                    ("FromStream".to_string(), Member::Str(from.to_string())),
                    ("ToStream".to_string(), Member::Str("Output".to_string())),
                ],
            ));
        } else {
            clean.push(tokens[i].clone());
        }
        i += 1;
    }
    (clean.join(" "), redirections)
}

fn redirection_operator(token: &str) -> Option<(&'static str, bool, Option<&str>)> {
    for (prefix, stream) in [
        ("*>", "All"),
        ("6>", "Information"),
        ("5>", "Debug"),
        ("4>", "Verbose"),
        ("3>", "Warning"),
        ("2>", "Error"),
        ("1>", "Output"),
        (">", "Output"),
    ] {
        if let Some(rest) = token.strip_prefix(prefix) {
            if rest.starts_with('&') {
                return None;
            }
            let append = rest.starts_with('>');
            let attached = if append { &rest[1..] } else { rest };
            return Some((stream, append, Some(attached)));
        }
    }
    None
}

fn merging_redirection(token: &str) -> Option<&'static str> {
    let (left, right) = token.split_once(">&")?;
    if right != "1" {
        return None;
    }
    Some(match left {
        "*" => "All",
        "6" => "Information",
        "5" => "Debug",
        "4" => "Verbose",
        "3" => "Warning",
        "2" => "Error",
        "1" | "" => "Output",
        _ => return None,
    })
}

fn command_element_nodes(text: &str, span: Span, nodes: &mut Vec<AstNode>) -> Vec<usize> {
    split_command_like_tokens(text)
        .into_iter()
        .map(|token| {
            if (token.starts_with('\'') && token.ends_with('\'')) || (token.starts_with('"') && token.ends_with('"')) {
                return source_expression_node(&token, span, nodes);
            }
            let kind = if token.starts_with('-') {
                "CommandParameterAst"
            } else {
                "StringConstantExpressionAst"
            };
            push(
                nodes,
                kind,
                span,
                Vec::new(),
                vec![
                    ("Value".to_string(), Member::Str(token.trim_start_matches('-').to_string())),
                    ("ParameterName".to_string(), Member::Str(token.trim_start_matches('-').to_string())),
                    ("StringConstantType".to_string(), Member::Str("BareWord".to_string())),
                    ("StaticType".to_string(), Member::Expr(type_ref_expr("System.String", false))),
                ],
            )
        })
        .collect()
}

fn source_looks_like_command(text: &str) -> bool {
    let Some(first) = text.split_whitespace().next() else {
        return false;
    };
    let lower = first.to_ascii_lowercase();
    if matches!(
        lower.as_str(),
        "if" | "else" | "elseif" | "for" | "foreach" | "while" | "do" | "switch" | "try" | "catch" | "finally"
    ) {
        return false;
    }
    let first_char = first.chars().next().unwrap_or('\0');
    first_char.is_ascii_alphabetic()
        && !text.contains('=')
        && !text.contains(" -eq ")
        && !text.contains(" -ne ")
        && !text.contains(" -gt ")
        && !text.contains(" -lt ")
        && !text.contains(" + ")
        && !text.contains(" * ")
        && !text.contains(" / ")
}

fn is_pipeline_expression_text(text: &str) -> bool {
    let trimmed = text.trim();
    if trimmed.is_empty() || trimmed.contains('=') {
        return false;
    }
    let lower = trimmed.to_ascii_lowercase();
    if [
        "if", "else", "elseif", "for", "foreach", "while", "do", "switch", "try", "catch", "finally", "function",
        "filter", "param",
    ]
    .into_iter()
    .any(|keyword| starts_keyword(&lower, keyword))
    {
        return false;
    }
    trimmed.starts_with('$')
        || trimmed.starts_with('\'')
        || trimmed.starts_with('"')
        || trimmed.starts_with("@(")
        || trimmed.starts_with("@{")
        || trimmed.starts_with('[')
        || trimmed.chars().next().is_some_and(|ch| ch.is_ascii_digit())
}

fn split_command_like_tokens(text: &str) -> Vec<String> {
    let mut tokens = Vec::new();
    let mut current = String::new();
    let mut quote: Option<char> = None;
    for ch in text.chars() {
        if let Some(q) = quote {
            current.push(ch);
            if ch == q {
                quote = None;
            }
            continue;
        }
        match ch {
            '\'' | '"' => {
                quote = Some(ch);
                current.push(ch);
            }
            c if c.is_whitespace() => {
                if !current.trim().is_empty() {
                    tokens.push(current.trim().to_string());
                    current.clear();
                }
            }
            _ => current.push(ch),
        }
    }
    if !current.trim().is_empty() {
        tokens.push(current.trim().to_string());
    }
    tokens
}

fn param(name: &str) -> Param {
    Param {
        name: name.to_string(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: true,
        is_nullable: false,
    }
}

fn lambda_expr(params: Vec<&str>, body: Vec<Statement>) -> Expression {
    Expression::new(ExprKind::Lambda {
        params: params.into_iter().map(param).collect(),
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    })
}

fn function_expr(params: Vec<&str>, body: Vec<Statement>) -> Expression {
    Expression::new(ExprKind::FunctionExpr(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name: String::new(),
            params: params.into_iter().map(param).collect(),
            body,
            modifiers: Modifiers::default(),
            is_async: false,
            is_generator: false,
            is_sub: false,
            return_type: None,
            handles: Vec::new(),
        },
    ))))
}

fn call_expr(callee: Expression, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn member_access(object: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(object),
        field: field.to_string(),
        null_safe: false,
    })
}

fn ast_method_find_all(owner: &str) -> Expression {
    let pred = format!("{owner}_findall_predicate");
    let nested = format!("{owner}_findall_nested");
    let node = format!("{owner}_findall_node");
    let box_name = format!("{owner}_findall_box");
    function_expr(
        vec![&pred, &nested],
        vec![
            Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(box_name.clone()),
                    type_hint: None,
                    init: Some(Expression::new(ExprKind::Object(vec![prop("Items", array_of(Vec::new()))]))),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::FunctionScoped,
            }),
            Statement::new(StmtKind::ForIn {
                var: node.clone(),
                key: None,
                iter: member_access(Expression::new(ExprKind::This), "__Descendants"),
                body: vec![Statement::new(StmtKind::If {
                    cond: call_expr(Expression::ident(&pred), vec![Expression::ident(&node)]),
                    then_body: vec![Statement::new(StmtKind::Assign {
                        targets: vec![member_access(Expression::ident(&box_name), "Items")],
                        value: array_append(
                            member_access(Expression::ident(&box_name), "Items"),
                            Expression::ident(&node),
                        ),
                        by_ref: false,
                    })],
                    elifs: Vec::new(),
                    else_body: None,
                })],
                of: true,
                else_body: None,
                is_async: false,
            }),
            Statement::new(StmtKind::Return(Some(member_access(Expression::ident(&box_name), "Items")))),
        ],
    )
}

fn array_append(array: Expression, value: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: vybe_ast::BinOp::Add,
        left: Box::new(array),
        right: Box::new(value),
    })
}

fn ast_method_find(owner: &str) -> Expression {
    let pred = format!("{owner}_find_predicate");
    let nested = format!("{owner}_find_nested");
    let node = format!("{owner}_find_node");
    function_expr(
        vec![&pred, &nested],
        vec![
            Statement::new(StmtKind::ForIn {
                var: node.clone(),
                key: None,
                iter: member_access(Expression::new(ExprKind::This), "__Descendants"),
                body: vec![Statement::new(StmtKind::If {
                    cond: call_expr(Expression::ident(&pred), vec![Expression::ident(&node)]),
                    then_body: vec![Statement::new(StmtKind::Return(Some(Expression::ident(&node))))],
                    elifs: Vec::new(),
                    else_body: None,
                })],
                of: true,
                else_body: None,
                is_async: false,
            }),
            Statement::new(StmtKind::Return(Some(Expression::null()))),
        ],
    )
}

fn ast_method_get_command_name(name: &str) -> Expression {
    function_expr(
        Vec::new(),
        vec![Statement::new(StmtKind::Return(Some(Expression::string(name))))],
    )
}

fn member_expr(member: &Member) -> Expression {
    match member {
        Member::Node(index) => Expression::ident(&temp(*index)),
        Member::Nodes(list) => array_of(list.iter().map(|i| Expression::ident(&temp(*i))).collect()),
        Member::Expr(expr) => expr.clone(),
        Member::Str(text) => Expression::string(text),
        Member::Int(value) => Expression::int(*value),
        Member::Bool(value) => Expression::bool(*value),
        Member::Null => Expression::null(),
    }
}

fn function_body_expr(text: &str) -> Expression {
    Expression::new(ExprKind::Object(vec![prop(
        "ParamBlock",
        Expression::new(ExprKind::Object(vec![prop(
            "Attributes",
            array_of(attribute_ast_exprs(text)),
        )])),
    )]))
}

fn attribute_ast_exprs(text: &str) -> Vec<Expression> {
    let Some(args) = attribute_args(text, "CmdletBinding") else {
        return Vec::new();
    };
    vec![Expression::new(ExprKind::Object(vec![
        prop("__type", Expression::string("AttributeAst")),
        prop("TypeName", Expression::new(ExprKind::Object(vec![prop(
            "Name",
            Expression::string("CmdletBinding"),
        )]))),
        prop(
            "NamedArguments",
            array_of(
                split_top_level_args(&args)
                    .into_iter()
                    .filter_map(|arg| {
                        let (name, value) = arg.split_once('=')?;
                        Some(named_attribute_argument_expr(name.trim(), value.trim()))
                    })
                    .collect(),
            ),
        ),
    ]))]
}

fn named_attribute_argument_expr(name: &str, value: &str) -> Expression {
    Expression::new(ExprKind::Object(vec![
        prop("__type", Expression::string("NamedAttributeArgumentAst")),
        prop("ArgumentName", Expression::string(name)),
        prop("Argument", attribute_argument_expr(value)),
    ]))
}

fn attribute_argument_expr(value: &str) -> Expression {
    let path = value.trim().trim_start_matches('$').trim();
    if path.eq_ignore_ascii_case("true") || path.eq_ignore_ascii_case("false") {
        return Expression::new(ExprKind::Object(vec![
            prop("__type", Expression::string("VariableExpressionAst")),
            prop(
                "VariablePath",
                Expression::new(ExprKind::Object(vec![prop(
                    "UserPath",
                    Expression::string(&path.to_ascii_lowercase()),
                )])),
            ),
        ]));
    }
    Expression::new(ExprKind::Object(vec![
        prop("__type", Expression::string("ConstantExpressionAst")),
        prop(
            "Value",
            Expression::string(value.trim().trim_matches('"').trim_matches('\'')),
        ),
    ]))
}

fn attribute_args(source: &str, attr: &str) -> Option<String> {
    let lower = source.to_ascii_lowercase();
    let needle = format!("[{}", attr.to_ascii_lowercase());
    let hit = lower.find(&needle)?;
    let mut rest = source[hit + needle.len()..].trim_start();
    if rest.starts_with(']') {
        return Some(String::new());
    }
    if !rest.starts_with('(') {
        return None;
    }
    rest = &rest[1..];
    let mut depth = 1i32;
    let mut quote: Option<char> = None;
    let mut end = 0usize;
    for (offset, ch) in rest.char_indices() {
        if let Some(q) = quote {
            if ch == q {
                quote = None;
            }
            continue;
        }
        match ch {
            '\'' | '"' => quote = Some(ch),
            '(' => depth += 1,
            ')' => {
                depth -= 1;
                if depth == 0 {
                    end = offset;
                    break;
                }
            }
            _ => {}
        }
    }
    Some(rest[..end].to_string())
}

fn split_top_level_args(text: &str) -> Vec<String> {
    let mut out = Vec::new();
    let mut current = String::new();
    let mut square = 0i32;
    let mut paren = 0i32;
    let mut quote: Option<char> = None;
    for ch in text.chars() {
        if let Some(q) = quote {
            current.push(ch);
            if ch == q {
                quote = None;
            }
            continue;
        }
        match ch {
            '\'' | '"' => {
                quote = Some(ch);
                current.push(ch);
            }
            '[' => {
                square += 1;
                current.push(ch);
            }
            ']' => {
                square -= 1;
                current.push(ch);
            }
            '(' => {
                paren += 1;
                current.push(ch);
            }
            ')' => {
                paren -= 1;
                current.push(ch);
            }
            ',' if square == 0 && paren == 0 => {
                out.push(current.trim().to_string());
                current.clear();
            }
            _ => current.push(ch),
        }
    }
    if !current.trim().is_empty() {
        out.push(current.trim().to_string());
    }
    out
}

fn extent_object(node: &AstNode, map: &SourceMap, whole: &str) -> Expression {
    let text = node
        .members
        .iter()
        .find_map(|(name, member)| match (name.as_str(), member) {
            ("__WholeText", Member::Str(text)) => Some(text.clone()),
            _ => None,
        })
        .unwrap_or_else(|| map.text(node.span).to_string());
    let start = map.offset(node.span.start_line, node.span.start_col);
    let end = if text == whole {
        whole.len()
    } else {
        start + text.len()
    };
    Expression::new(ExprKind::Object(vec![
        prop("Text", Expression::string(&text)),
        prop("StartOffset", Expression::int(start as i64)),
        prop("EndOffset", Expression::int(end as i64)),
        prop(
            "StartLineNumber",
            Expression::int(node.span.start_line.max(1) as i64),
        ),
        prop(
            "StartColumnNumber",
            Expression::int(node.span.start_col.max(1) as i64),
        ),
        prop(
            "EndLineNumber",
            Expression::int(node.span.end_line.max(1) as i64),
        ),
        prop(
            "EndColumnNumber",
            Expression::int(node.span.end_col.max(1) as i64),
        ),
        prop("File", Expression::null()),
    ]))
}

fn descendants(nodes: &[AstNode], root: usize, out: &mut Vec<usize>) {
    out.push(root);
    for &child in &nodes[root].children {
        descendants(nodes, child, out);
    }
}

fn emit_tree(nodes: &[AstNode], root: usize, map: &SourceMap, source: &str) -> Expression {
    let mut body: Vec<Statement> = Vec::new();

    for (index, node) in nodes.iter().enumerate() {
        let mut props = vec![
            prop("__type", Expression::string(&node.ty)),
            prop("__types", array_of(ast_type_names(&node.ty).into_iter().map(|name| Expression::string(&name)).collect())),
            prop("Extent", extent_object(node, map, source)),
            prop(
                "Children",
                array_of(
                    node.children
                        .iter()
                        .map(|i| Expression::ident(&temp(*i)))
                        .collect(),
                ),
            ),
            prop("Parent", Expression::null()),
            prop("FindAll", ast_method_find_all(&temp(index))),
            prop("Find", ast_method_find(&temp(index))),
        ];
        if node.ty == "CommandAst" {
            let command_name = node
                .members
                .iter()
                .find_map(|(name, member)| match (name.as_str(), member) {
                    ("CommandName", Member::Str(value)) => Some(value.as_str()),
                    _ => None,
                })
                .unwrap_or("");
            props.push(prop("GetCommandName", ast_method_get_command_name(command_name)));
        }
        for (name, member) in &node.members {
            if name == "__WholeText" {
                continue;
            }
            props.push(prop(name, member_expr(member)));
        }
        body.push(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(temp(index)),
                type_hint: None,
                init: Some(Expression::new(ExprKind::Object(props))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::FunctionScoped,
        }));
    }

    // ⛔`Parent` AND THE DESCENDANT LISTS CLOSE CYCLES, so they are written once
    // every node exists. An object literal is constructed WHERE IT APPEARS, so a
    // node named in two literals would be two objects at run time and `Find`
    // would answer a node whose members are not the ones the walk reached.
    for (index, node) in nodes.iter().enumerate() {
        for &child in &node.children {
            body.push(assign_member(
                &temp(child),
                "Parent",
                Expression::ident(&temp(index)),
            ));
        }
        let mut reachable = Vec::new();
        descendants(nodes, index, &mut reachable);
        body.push(assign_member(
            &temp(index),
            "__Descendants",
            array_of(
                reachable
                    .iter()
                    .map(|i| Expression::ident(&temp(*i)))
                    .collect(),
            ),
        ));
    }

    body.push(Statement::new(StmtKind::Return(Some(Expression::ident(
        &temp(root),
    )))));

    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: "_".to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: true,
                is_nullable: false,
            }],
            body: LambdaBody::Block(body),
            is_async: false,
            captures: Vec::new(),
        })),
        args: Vec::new(),
        optional: false,
    })
}

fn ast_type_names(ty: &str) -> Vec<String> {
    let mut names = vec![ty.to_string(), "Ast".to_string()];
    if ty.ends_with("ExpressionAst") || matches!(ty, "CommandAst" | "PipelineAst" | "HashtableAst" | "ArrayLiteralAst") {
        names.push("ExpressionAst".to_string());
    }
    if ty == "TypeConstraintAst" {
        names.push("AttributeBaseAst".to_string());
    }
    if ty.ends_with("StatementAst") {
        names.push("StatementAst".to_string());
    }
    names
}

fn assign_member(target: &str, field: &str, value: Expression) -> Statement {
    Statement::new(StmtKind::Assign {
        targets: vec![Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(target)),
            field: field.to_string(),
            null_safe: false,
        })],
        value,
        by_ref: false,
    })
}
