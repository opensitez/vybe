//! Hoist backing selection out of proven forward byte-copy loops.

#[path = "normalize_affine_words.rs"]
mod words;

use std::collections::{HashMap, HashSet};
use vybe_ast::{
    Argument, BinOp, BindingPattern, BitLane, ExprKind, Expression, IntOp, Literal, Param, PassBy,
    PlaceExpr, Statement, StmtKind, UnaryOp,
};

fn ident(value: &Expression) -> Option<&str> {
    if let ExprKind::Ident(name) = &value.kind {
        Some(name)
    } else {
        None
    }
}

fn int(value: i64) -> Expression {
    Expression::int(value)
}
fn binary(op: BinOp, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}
fn word(op: IntOp, lane: BitLane, left: Expression, right: Expression) -> Expression {
    binary(BinOp::Integer(op, lane), left, right)
}
fn call(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}
fn assign(name: &str, value: Expression) -> Statement {
    Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
        target: Box::new(Expression::ident(name)),
        value: Box::new(value),
    })))
}
fn number(name: &str) -> Expression {
    Expression::new(ExprKind::IsType {
        expr: Box::new(Expression::ident(name)),
        type_name: "number".into(),
    })
}
fn unsigned(name: &str) -> Expression {
    Expression::new(ExprKind::Cast {
        expr: Box::new(Expression::ident(name)),
        type_name: "uint32".into(),
    })
}

#[derive(Default)]
struct Context {
    declared: HashSet<String>,
    local: HashSet<String>,
    integers: HashSet<String>,
    ordinary: HashSet<String>,
    forbidden: HashSet<String>,
    references: HashMap<String, usize>,
    nested_function: bool,
}

fn integer_hint(hint: Option<&vybe_ast::TypeHint>) -> bool {
    hint.is_some_and(|hint| {
        hint.binding != vybe_ast::TypeBinding::Descriptive
            && matches!(hint.spelling().trim(), "int" | "int32" | "integer")
    })
}

pub(crate) fn children(statement: &mut Statement, visit: &mut impl FnMut(&mut Vec<Statement>)) {
    match &mut statement.kind {
        StmtKind::Block(body)
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. }
        | StmtKind::For { body, .. }
        | StmtKind::NamespaceDecl { body, .. } => visit(body),
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            visit(then_body);
            for (_, body) in elifs {
                visit(body);
            }
            if let Some(body) = else_body {
                visit(body);
            }
        }
        StmtKind::Switch { cases, default, .. } => {
            for case in cases {
                visit(&mut case.body);
            }
            if let Some(body) = default {
                visit(body);
            }
        }
        _ => {}
    }
}

fn declarations(body: &mut Vec<Statement>, context: &mut Context) {
    for statement in body {
        context.nested_function |= matches!(statement.kind, StmtKind::FunctionDecl { .. });
        if let StmtKind::VarDecl { declarations, .. } = &statement.kind {
            for declaration in declarations {
                if let BindingPattern::Ident(name) = &declaration.pattern {
                    // Ambiguous/shadowed declarations cannot supply a type proof.
                    if !context.declared.insert(name.clone()) {
                        context.forbidden.insert(name.clone());
                    }
                    if integer_hint(declaration.type_hint.as_ref()) {
                        context.integers.insert(name.clone());
                    }
                    if declaration
                        .type_hint
                        .as_ref()
                        .is_some_and(|hint| hint.access == Some(vybe_ast::StorageAccess::Ordinary))
                    {
                        context.ordinary.insert(name.clone());
                    }
                }
            }
        }
        children(statement, &mut |body| declarations(body, context));
    }
}

fn inspect(value: &Expression, context: &mut Context) {
    if let Some(name) = ident(value) {
        *context.references.entry(name.into()).or_default() += 1;
    }
    match &value.kind {
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => {
            if let Some(name) = ident(expr) {
                context.forbidden.insert(name.into());
            }
        }
        ExprKind::RefOf(place) => {
            if let PlaceExpr::Ident(name) = &**place {
                context.forbidden.insert(name.clone());
            }
        }
        ExprKind::Lambda { captures, .. } => {
            // Captured bindings may be observed independently of this loop.
            if !captures.is_empty() {
                context.forbidden.extend(context.declared.iter().cloned());
            }
        }
        _ => {}
    }
}

fn context(body: &mut Vec<Statement>, params: &[Param], globals: &HashSet<String>) -> Context {
    let mut context = Context::default();
    for param in params {
        context.declared.insert(param.name.clone());
        if integer_hint(param.type_hint.as_ref()) {
            context.integers.insert(param.name.clone());
        }
        if param
            .type_hint
            .as_ref()
            .is_some_and(|hint| hint.access == Some(vybe_ast::StorageAccess::Ordinary))
        {
            context.ordinary.insert(param.name.clone());
        }
        if param.pass_by != PassBy::Value {
            context.forbidden.insert(param.name.clone());
        }
    }
    declarations(body, &mut context);
    context.local = context.declared.clone();
    context.declared.extend(globals.iter().cloned());
    if context.nested_function {
        context.forbidden.extend(context.local.iter().cloned());
    }
    for statement in body {
        statement.walk_exprs(&mut |value| inspect(value, &mut context));
    }
    context
}

#[derive(Clone, Debug, PartialEq)]
enum Value {
    Cursor(String, i64),
    Byte(String, i64),
    Integer(i64),
    Boolean(bool),
}

struct Proof {
    values: HashMap<String, Value>,
    writes: HashSet<String>,
    reads: Vec<(String, i64)>,
    stores: Vec<(Value, Value)>,
}

impl Proof {
    fn evaluate(&mut self, value: &Expression) -> Option<Value> {
        match &value.kind {
            ExprKind::Ident(name) => self.values.get(name).cloned(),
            ExprKind::Lit(Literal::Int(n)) => Some(Value::Integer(*n)),
            ExprKind::Lit(Literal::Bool(value)) => Some(Value::Boolean(*value)),
            ExprKind::Unary {
                op: UnaryOp::Neg,
                expr,
            } => {
                let Value::Integer(value) = self.evaluate(expr)? else {
                    return None;
                };
                Some(Value::Integer(value.checked_neg()?))
            }
            ExprKind::Sequence(values) => {
                let mut result = None;
                for value in values {
                    result = Some(self.evaluate(value)?);
                }
                result
            }
            ExprKind::Assign { target, value } => {
                let name = ident(target)?.to_owned();
                let result = self.evaluate(value)?;
                self.writes.insert(name.clone());
                self.values.insert(name, result.clone());
                Some(result)
            }
            ExprKind::IsType { expr, type_name } if type_name == "number" => Some(Value::Boolean(
                !matches!(self.evaluate(expr)?, Value::Boolean(_)),
            )),
            ExprKind::Ternary { cond, then, else_ } => match self.evaluate(cond)? {
                Value::Boolean(true) => self.evaluate(then),
                Value::Boolean(false) => self.evaluate(else_),
                _ => None,
            },
            ExprKind::Binary {
                op: BinOp::Or | BinOp::And,
                left,
                right,
            } => {
                let Value::Boolean(left) = self.evaluate(left)? else {
                    return None;
                };
                if matches!(&value.kind, ExprKind::Binary { op: BinOp::Or, .. }) == left {
                    Some(Value::Boolean(left))
                } else {
                    self.evaluate(right)
                }
            }
            ExprKind::Binary {
                op: BinOp::Add | BinOp::Integer(IntOp::Add, BitLane::W32),
                left,
                right,
            } => match (self.evaluate(left)?, self.evaluate(right)?) {
                (Value::Cursor(name, offset), Value::Integer(step))
                    if (i32::MIN as i64..=i32::MAX as i64).contains(&step) =>
                {
                    Some(Value::Cursor(name, offset.checked_add(step)?))
                }
                _ => None,
            },
            ExprKind::Binary {
                op: BinOp::BitAnd | BinOp::Integer(IntOp::And, BitLane::W32),
                left,
                right,
            } => match (self.evaluate(left)?, self.evaluate(right)?) {
                (byte @ Value::Byte(..), Value::Integer(255)) => Some(byte),
                _ => None,
            },
            ExprKind::Cast { expr, type_name } if !type_name.contains("volatile") => {
                match self.evaluate(expr)? {
                    cursor @ Value::Cursor(..)
                        if type_name.contains('*') || type_name == "pointerCast" =>
                    {
                        Some(cursor)
                    }
                    byte @ Value::Byte(..) if type_name == "uint8" => Some(byte),
                    _ => None,
                }
            }
            ExprKind::Call {
                callee,
                args,
                optional: false,
            } if args
                .iter()
                .all(|arg| !arg.by_ref && !arg.spread && arg.name.is_none()) =>
            {
                match (ident(callee)?, args.as_slice()) {
                    ("__c_ptr_i32_load8_u", [address]) => {
                        let Value::Cursor(name, offset) = self.evaluate(&address.value)? else {
                            return None;
                        };
                        self.reads.push((name.clone(), offset));
                        Some(Value::Byte(name, offset))
                    }
                    ("__c_ptr_i32_store8" | "__c_ptr_byte_store", [address, byte]) => {
                        let address = self.evaluate(&address.value)?;
                        let byte = self.evaluate(&byte.value)?;
                        self.stores.push((address, byte.clone()));
                        Some(byte)
                    }
                    _ => None,
                }
            }
            _ => None,
        }
    }
}

fn countdown(value: &Expression) -> Option<&str> {
    match &value.kind {
        ExprKind::Unary {
            op: UnaryOp::PostDec,
            expr,
        } => ident(expr),
        ExprKind::Binary {
            op:
                BinOp::NotEq
                | BinOp::Integer(IntOp::Ne, BitLane::W32)
                | BinOp::Integer(IntOp::Or, BitLane::W32),
            left,
            right,
        } if matches!(right.kind, ExprKind::Lit(Literal::Int(0))) => countdown(left),
        _ => None,
    }
}

fn match_loop(statement: &Statement, context: &Context) -> Option<(String, String, String, i64)> {
    let StmtKind::While {
        cond,
        body,
        else_body: None,
    } = &statement.kind
    else {
        return None;
    };
    let count = countdown(cond)?.to_owned();
    if !context.integers.contains(&count) || body.len() != 2 {
        return None;
    }
    let StmtKind::Expr(update) = &body[1].kind else {
        return None;
    };
    let ExprKind::Assign { target, .. } = &update.kind else {
        return None;
    };
    let destination = ident(target)?.to_owned();
    let mut sources = HashSet::new();
    body[0].walk_exprs(&mut |value| {
        if let ExprKind::Assign { target, .. } = &value.kind {
            if let Some(name) = ident(target) {
                if name != destination && name != count && context.local.contains(name) {
                    sources.insert(name.to_owned());
                }
            }
        }
    });
    if sources.len() != 1 {
        return None;
    }
    let source = sources.into_iter().next()?;
    if destination == count
        || [&source, &destination, &count].iter().any(|name| {
            !context.local.contains(*name)
                || !context.ordinary.contains(*name)
                || context.forbidden.contains(*name)
        })
    {
        return None;
    }
    let mut proof = Proof {
        values: [source.clone(), destination.clone()]
            .into_iter()
            .map(|name| (name.clone(), Value::Cursor(name, 0)))
            .collect(),
        writes: HashSet::new(),
        reads: vec![],
        stores: vec![],
    };
    for statement in body {
        let StmtKind::Expr(value) = &statement.kind else {
            return None;
        };
        proof.evaluate(value)?;
    }
    if proof.reads != [(source.clone(), 0)]
        || proof.stores
            != [(
                Value::Cursor(destination.clone(), 0),
                Value::Byte(source.clone(), 0),
            )]
        || proof.values.get(&source) != Some(&Value::Cursor(source.clone(), 1))
    {
        return None;
    }
    let Value::Cursor(base, stride) = proof.values.get(&destination)? else {
        return None;
    };
    if base != &destination || !(i32::MIN as i64..=i32::MAX as i64).contains(stride) {
        return None;
    }
    let mut inside = Context::default();
    statement.walk_exprs(&mut |value| inspect(value, &mut inside));
    for name in &proof.writes {
        if name == &source || name == &destination {
            continue;
        }
        if context.declared.contains(name)
            || context.references.get(name) != inside.references.get(name)
        {
            return None;
        }
    }
    Some((destination, source, count, *stride))
}

fn lower(statement: &mut Statement, destination: &str, source: &str, count: &str, stride: i64) {
    let n = || Expression::ident(count);
    let memory_bytes = || {
        word(
            IntOp::Mul,
            BitLane::W64,
            call("__c_ptr_memory_size", vec![]),
            int(65536),
        )
    };
    let last = word(
        IntOp::Add,
        BitLane::W64,
        unsigned(destination),
        word(
            IntOp::Mul,
            BitLane::W64,
            word(IntOp::Sub, BitLane::W64, n(), int(1)),
            int(stride),
        ),
    );
    let source_end = word(IntOp::Add, BitLane::W64, unsigned(source), n());
    let destination_end = word(
        IntOp::Add,
        BitLane::W64,
        unsigned(destination),
        word(IntOp::Mul, BitLane::W64, n(), int(stride)),
    );
    let numeric_address = |name: &str| {
        Expression::new(ExprKind::Cast {
            expr: Box::new(Expression::ident(name)),
            type_name: "double".into(),
        })
    };
    let bounds = [
        // Plain numeric Add and word Add coincide throughout this domain,
        // including the final one-past cursors observed after the loop.
        binary(BinOp::Eq, numeric_address(source), unsigned(source)),
        binary(
            BinOp::Eq,
            numeric_address(destination),
            unsigned(destination),
        ),
        binary(BinOp::GtEq, numeric_address(source), int(0)),
        binary(BinOp::GtEq, numeric_address(destination), int(0)),
        binary(BinOp::LtEq, numeric_address(source), int(i32::MAX as i64)),
        binary(
            BinOp::LtEq,
            numeric_address(destination),
            int(i32::MAX as i64),
        ),
        word(IntOp::LeS, BitLane::W64, source_end, int(i32::MAX as i64)),
        word(IntOp::GeS, BitLane::W64, destination_end.clone(), int(0)),
        word(
            IntOp::LeS,
            BitLane::W64,
            destination_end,
            int(i32::MAX as i64),
        ),
        word(
            IntOp::LeU,
            BitLane::W64,
            word(IntOp::Add, BitLane::W64, unsigned(source), n()),
            memory_bytes(),
        ),
        word(IntOp::GeS, BitLane::W64, last.clone(), int(0)),
        word(IntOp::LtU, BitLane::W64, last, memory_bytes()),
        word(
            IntOp::LtU,
            BitLane::W64,
            unsigned(destination),
            memory_bytes(),
        ),
    ]
    .into_iter()
    .reduce(|left, right| binary(BinOp::And, left, right))
    .unwrap();
    let types = [destination, source, count]
        .into_iter()
        .map(number)
        .reduce(|left, right| binary(BinOp::And, left, right))
        .unwrap();
    let condition = binary(
        BinOp::And,
        types,
        binary(
            BinOp::And,
            word(IntOp::GeS, BitLane::W32, n(), int(0)),
            binary(
                BinOp::Or,
                word(IntOp::Eq, BitLane::W32, n(), int(0)),
                bounds,
            ),
        ),
    );
    let transfer = vec![
        Statement::new(StmtKind::Expr(call(
            "__c_ptr_copy_column",
            vec![
                Expression::ident(destination),
                Expression::ident(source),
                int(stride),
                n(),
            ],
        ))),
        assign(
            source,
            word(IntOp::Add, BitLane::W32, Expression::ident(source), n()),
        ),
        assign(
            destination,
            word(
                IntOp::Add,
                BitLane::W32,
                Expression::ident(destination),
                word(IntOp::Mul, BitLane::W32, n(), int(stride)),
            ),
        ),
    ];
    let scalar = statement.clone();
    statement.kind = StmtKind::If {
        cond: condition,
        then_body: vec![
            Statement::new(StmtKind::If {
                cond: word(IntOp::Ne, BitLane::W32, n(), int(0)),
                then_body: transfer,
                elifs: vec![],
                else_body: None,
            }),
            assign(count, int(-1)),
        ],
        elifs: vec![],
        else_body: Some(vec![scalar]),
    };
}

fn loops(body: &mut Vec<Statement>, context: &Context) {
    for statement in body {
        if let StmtKind::If {
            then_body,
            else_body: Some(fallback),
            ..
        } = &statement.kind
        {
            if fallback.len() == 1 && matches!(fallback[0].kind, StmtKind::While { .. } | StmtKind::For { .. }) {
                let mut transferred = false;
                for statement in then_body {
                    statement.walk_exprs(&mut |value| {
                        transferred |= matches!(&value.kind, ExprKind::Call { callee, .. }
                            if matches!(ident(callee), Some("__c_ptr_copy_column" | "__c_ptr_copy_column16")));
                    });
                }
                if transferred {
                    continue;
                }
            }
        }
        if words::normalize_for(statement, context) {
            continue;
        } else if let Some((destination, source, count, stride)) = match_loop(statement, context) {
            lower(statement, &destination, &source, &count, stride);
        } else {
            children(statement, &mut |body| loops(body, context));
        }
    }
}

pub fn normalize(body: &mut Vec<Statement>) {
    let globals = body
        .iter()
        .filter_map(|statement| {
            if let StmtKind::VarDecl { declarations, .. } = &statement.kind {
                Some(declarations.iter().filter_map(|declaration| {
                    if let BindingPattern::Ident(name) = &declaration.pattern {
                        Some(name.clone())
                    } else {
                        None
                    }
                }))
            } else {
                None
            }
        })
        .flatten()
        .collect();
    for statement in body {
        if let StmtKind::FunctionDecl { params, body, .. } = &mut statement.kind {
            let context = context(body, params, &globals);
            loops(body, &context);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn transfers(source: &str) -> usize {
        let mut module = crate::parse(source).unwrap();
        normalize(&mut module.body);
        let mut count = 0;
        for statement in &module.body {
            statement.walk_exprs(&mut |value| {
                if matches!(&value.kind, ExprKind::Call { callee, .. }
                    if ident(callee) == Some("__c_ptr_copy_column"))
                {
                    count += 1;
                }
            });
        }
        count
    }

    #[test]
    fn connects_generic_typed_affine_loop_and_typedef_local_cursors() {
        for source in [
            "void scatter(unsigned char *d, unsigned char *s, int n) {
                while (n--) { *d = *s++; d += 3; } }",
            "typedef unsigned char octet;
             void arbitrary(void *out, void *in, int length) {
                octet *d = out; octet *s = in; int n = length;
                while (n--) { *d = *s++; d += 320; } }",
        ] {
            assert_eq!(transfers(source), 1, "{source}");
        }
    }

    #[test]
    fn preserves_unproved_widths_effects_and_exposed_bindings() {
        for source in [
            "void f(unsigned short *d, unsigned short *s, int n) {
                while (n--) { *d = *s++; d += 3; } }",
            "void f(unsigned char *d, unsigned char *s, unsigned int n) {
                while (n--) { *d = *s++; d += 3; } }",
            "void f(unsigned char *d, unsigned char *s, int n) {
                int *alias = &n; while (n--) { *d = *s++; d += 3; } }",
            "void f(unsigned char *d, unsigned char *s, int n) {
                unsigned char **alias = &d; while (n--) { *d = *s++; d += 3; } }",
            "void f(unsigned char *d, unsigned char *s, int n) {
                unsigned char **alias = &s; while (n--) { *d = *s++; d += 3; } }",
            "void f(unsigned char *d, unsigned char *s, int n) {
                while (n--) { *d = *s++; d += 3; n--; } }",
            "void f(volatile unsigned char *d, unsigned char *s, int n) {
                while (n--) { *d = *s++; d += 3; } }",
            "void f(unsigned char *d, volatile unsigned char *s, int n) {
                while (n--) { *d = *s++; d += 3; } }",
            "void f(unsigned char *d, unsigned char *s, volatile int n) {
                while (n--) { *d = *s++; d += 3; } }",
            "typedef volatile int counter;
             void f(unsigned char *d, unsigned char *s, counter n) {
                while (n--) { *d = *s++; d += 3; } }",
        ] {
            assert_eq!(transfers(source), 0, "{source}");
        }
    }

    #[test]
    fn retains_scalar_fallback_and_does_not_rewrite_it_again() {
        let mut module = crate::parse(
            "void f(unsigned char *d, unsigned char *s, int n) {
            while (n--) { *d = *s++; d += 3; } }",
        )
        .unwrap();
        normalize(&mut module.body);
        let before = format!("{:?}", module.body);
        normalize(&mut module.body);
        assert_eq!(format!("{:?}", module.body), before);
        assert!(before.contains("__c_ptr_memory_size"));
        assert!(before.contains("While"));
        assert!(before.contains("PostDec"));
    }
}
