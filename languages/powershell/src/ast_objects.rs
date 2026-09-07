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
    ArrayElement, BindingPattern, ExprKind, Expression, LambdaBody, Literal, Module,
    ObjectProperty, Param, PassBy, Span, Statement, StmtKind, VarDeclKind, VarDeclarator,
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
    for stmt in &module.body {
        if let Some(index) = walk_stmt(stmt, &map, &mut nodes) {
            roots.push(index);
        }
    }
    let root = push(
        &mut nodes,
        "ScriptBlockAst",
        Span {
            start_line: 1,
            start_col: 1,
            end_line: map.line_starts.len() as u32,
            end_col: u32::MAX,
        },
        roots,
        Vec::new(),
    );
    // The root's extent is the whole script, whatever the last line's length.
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

// ── translation ───────────────────────────────────────────────────────────

fn walk_stmt(stmt: &Statement, map: &SourceMap, nodes: &mut Vec<AstNode>) -> Option<usize> {
    let text = map.text(stmt.span).trim_start();
    match &stmt.kind {
        StmtKind::Assign { targets, value, .. } => {
            let left = targets.first().and_then(|t| walk_expr(t, map, nodes));
            let right = walk_expr(value, map, nodes);
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
        StmtKind::Expr(expr) => walk_expr(expr, map, nodes),
        // ⛔`switch` NORMALIZES TO `If`. The lowering is right for execution and
        // wrong for inspection, and the span says which one the source spelled.
        StmtKind::If { .. } if text.starts_with("switch") => {
            let children = child_statements(stmt, map, nodes);
            let members = vec![("Flags".to_string(), Member::Str(switch_flags(text)))];
            Some(push(
                nodes,
                "SwitchStatementAst",
                stmt.span,
                children,
                members,
            ))
        }
        StmtKind::If { .. } => {
            let children = child_statements(stmt, map, nodes);
            Some(push(nodes, "IfStatementAst", stmt.span, children, Vec::new()))
        }
        StmtKind::While { .. } if text.starts_with("do") => {
            let children = child_statements(stmt, map, nodes);
            Some(push(
                nodes,
                "DoWhileStatementAst",
                stmt.span,
                children,
                Vec::new(),
            ))
        }
        StmtKind::While { .. } => {
            let children = child_statements(stmt, map, nodes);
            Some(push(
                nodes,
                "WhileStatementAst",
                stmt.span,
                children,
                Vec::new(),
            ))
        }
        StmtKind::For { .. } => {
            let children = child_statements(stmt, map, nodes);
            Some(push(
                nodes,
                "ForStatementAst",
                stmt.span,
                children,
                Vec::new(),
            ))
        }
        StmtKind::ForIn { .. } => {
            let children = child_statements(stmt, map, nodes);
            Some(push(
                nodes,
                "ForEachStatementAst",
                stmt.span,
                children,
                Vec::new(),
            ))
        }
        StmtKind::Try { .. } => {
            let children = child_statements(stmt, map, nodes);
            Some(push(
                nodes,
                "TryStatementAst",
                stmt.span,
                children,
                Vec::new(),
            ))
        }
        StmtKind::Throw { .. } => {
            let children = child_statements(stmt, map, nodes);
            Some(push(
                nodes,
                "ThrowStatementAst",
                stmt.span,
                children,
                Vec::new(),
            ))
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
        StmtKind::FunctionDecl { name, .. } => {
            let children = child_statements(stmt, map, nodes);
            let members = vec![
                ("Name".to_string(), Member::Str(name.clone())),
                ("IsFilter".to_string(), Member::Bool(text.starts_with("filter"))),
                ("Body".to_string(), Member::Expr(function_body_expr(text))),
            ];
            Some(push(
                nodes,
                "FunctionDefinitionAst",
                stmt.span,
                children,
                members,
            ))
        }
        StmtKind::VarDecl { declarations, .. } => {
            let mut children = Vec::new();
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
    match &expr.kind {
        ExprKind::Ident(name) => Some(push(
            nodes,
            "VariableExpressionAst",
            expr.span,
            Vec::new(),
            vec![(
                "VariablePath".to_string(),
                Member::Str(name.trim_start_matches('$').to_string()),
            )],
        )),
        // ⛔THE QUOTE IS NOT IN THE NODE. Both spellings are `Lit(Str)`; the
        // span's first character is what tells them apart, and PowerShell's
        // `StringConstantType` is exactly that distinction.
        ExprKind::Lit(Literal::Str(value)) => {
            let single = text.trim_start().starts_with('\'');
            let ty = if single {
                "StringConstantExpressionAst"
            } else {
                "ExpandableStringExpressionAst"
            };
            let kind = if single { "SingleQuoted" } else { "DoubleQuoted" };
            Some(push(
                nodes,
                ty,
                expr.span,
                Vec::new(),
                vec![
                    ("Value".to_string(), Member::Str(value.clone())),
                    ("StringConstantType".to_string(), Member::Str(kind.into())),
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
                vec![("TypeName".to_string(), Member::Str(type_name.clone()))],
            );
            let children: Vec<usize> = std::iter::once(ty).chain(child).collect();
            Some(push(
                nodes,
                "ConvertExpressionAst",
                expr.span,
                children,
                Vec::new(),
            ))
        }
        ExprKind::Array(items) => {
            let children = items
                .iter()
                .filter_map(|item| walk_expr(&item.value, map, nodes))
                .collect();
            Some(push(
                nodes,
                "ArrayLiteralAst",
                expr.span,
                children,
                Vec::new(),
            ))
        }
        ExprKind::Object(_) => Some(push(
            nodes,
            "HashtableAst",
            expr.span,
            Vec::new(),
            Vec::new(),
        )),
        ExprKind::Lambda { .. } => Some(push(
            nodes,
            "ScriptBlockExpressionAst",
            expr.span,
            Vec::new(),
            Vec::new(),
        )),
        // ⛔A PIPELINE IS FOLDED INTO METHOD CALLS. Execution wants the fold and
        // inspection wants the stages; the span says which the source spelled,
        // and its own `|` marks the boundaries.
        ExprKind::Call { callee, args, .. } => {
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
                vec![("PipelineElements".to_string(), Member::Nodes(children.clone()))]
            } else {
                vec![("CommandElements".to_string(), Member::Nodes(children.clone()))]
            };
            Some(push(nodes, ty, expr.span, children, members))
        }
        ExprKind::Member { object, .. } => walk_expr(object, map, nodes),
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

fn switch_flags(text: &str) -> String {
    let mut flags = Vec::new();
    for word in ["Regex", "Wildcard", "Exact", "CaseSensitive", "File", "Parallel"] {
        if text
            .to_lowercase()
            .contains(&format!("-{}", word.to_lowercase()))
        {
            flags.push(word.to_string());
        }
    }
    if flags.is_empty() {
        flags.push("None".into());
    }
    flags.join(", ")
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
        ];
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
