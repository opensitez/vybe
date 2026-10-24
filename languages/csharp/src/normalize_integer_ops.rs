//! C# integral arithmetic as typed integer operations.
//!
//! An operator whose operands have C# integral types states its width and
//! signedness on the node, `BinOp::Integer(op, lane)`, the same node every
//! statically typed frontend emits for a machine-word operation. The type is
//! the one C# itself assigns (ECMA-334 §12.4.7 binary numeric promotion):
//! `int op int` is `int`, an `int` meeting a `long` is `long`, `uint` meeting a
//! signed `int` is `long`. An operand whose type the declaration does not state
//! leaves its operator untyped, exactly as before.
//!
//! C# integer arithmetic is unchecked by default, which is what the wasm
//! integer instructions do: `int.MaxValue + 1` wraps. Division and remainder
//! are typed only by a constant divisor other than 0 and -1, because a zero
//! divisor must throw `DivideByZeroException` and `MinValue / -1` must throw
//! `OverflowException`, where the instructions trap.

use std::collections::HashMap;

use vybe_ast::{
    BinOp, BitLane, ClassMember, CompoundOp, ExprKind, Expression, IntOp, LambdaBody, Literal,
    Statement, StmtKind, UnaryOp,
};

/// A C# integral type, after promotion of the types narrower than `int`.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum IntTy {
    I32,
    U32,
    I64,
    U64,
}

impl IntTy {
    fn lane(self) -> BitLane {
        match self {
            IntTy::I32 | IntTy::U32 => BitLane::W32,
            IntTy::I64 | IntTy::U64 => BitLane::W64,
        }
    }

    fn signed(self) -> bool {
        matches!(self, IntTy::I32 | IntTy::I64)
    }

    fn spelling(self) -> &'static str {
        match self {
            IntTy::I32 => "int",
            IntTy::U32 => "uint",
            IntTy::I64 => "long",
            IntTy::U64 => "ulong",
        }
    }
}

/// A declared C# integral type: its promoted operand type, and whether storing
/// a promoted value back narrows it (`short`, `byte`, …).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
struct Declared {
    ty: IntTy,
    narrow: bool,
}

fn declared(spelling: &str) -> Option<Declared> {
    let spelling = spelling.trim();
    let spelling = spelling.strip_prefix("System.").unwrap_or(spelling);
    let (ty, narrow) = match spelling {
        "int" | "Int32" => (IntTy::I32, false),
        "uint" | "UInt32" => (IntTy::U32, false),
        "long" | "Int64" => (IntTy::I64, false),
        "ulong" | "UInt64" => (IntTy::U64, false),
        "short" | "ushort" | "byte" | "sbyte" | "Int16" | "UInt16" | "Byte" | "SByte" => {
            (IntTy::I32, true)
        }
        _ => return None,
    };
    Some(Declared { ty, narrow })
}

/// The element type spelled by an array or list type: `long[]`, `List<long>`.
fn element_spelling(spelling: &str) -> Option<&str> {
    let spelling = spelling.trim();
    if let Some(element) = spelling.strip_suffix("[]") {
        return Some(element.trim());
    }
    let spelling = spelling.strip_prefix("System.Collections.Generic.").unwrap_or(spelling);
    ["List<", "IList<", "IEnumerable<", "IReadOnlyList<", "ICollection<"]
        .iter()
        .find_map(|head| spelling.strip_prefix(head))
        .and_then(|rest| rest.strip_suffix('>'))
        .map(str::trim)
}

/// The declared spelling a narrowing store converts to.
fn narrow_spelling(spelling: &str) -> String {
    let spelling = spelling.trim();
    spelling.strip_prefix("System.").unwrap_or(spelling).to_owned()
}

/// Names whose C# integral type is known, for one method body. A name
/// declared twice with different types is left unknown.
#[derive(Clone, Default)]
struct Env {
    names: HashMap<String, Option<(Declared, String)>>,
    /// The integral element type of a name declared as an array or list.
    elements: HashMap<String, Option<(Declared, String)>>,
    methods: HashMap<String, Option<IntTy>>,
    /// The declared integral return type of the enclosing method.
    returns: Option<IntTy>,
}

impl Env {
    fn declare(&mut self, name: &str, spelling: Option<&str>) {
        let entry = spelling.and_then(|s| declared(s).map(|d| (d, narrow_spelling(s))));
        let element = spelling.and_then(element_spelling)
            .and_then(|s| declared(s).map(|d| (d, narrow_spelling(s))));
        for (map, entry) in [(&mut self.names, entry), (&mut self.elements, element)] {
            match map.get(name) {
                Some(existing) if existing.as_ref().map(|e| e.0) != entry.as_ref().map(|e| e.0) => {
                    map.insert(name.to_owned(), None);
                }
                Some(_) => {}
                None => {
                    map.insert(name.to_owned(), entry);
                }
            }
        }
    }

    fn declare_method(&mut self, name: &str, return_type: Option<&str>) {
        let ty = return_type.and_then(declared).map(|d| d.ty);
        match self.methods.get(name) {
            Some(existing) if *existing != ty => {
                self.methods.insert(name.to_owned(), None);
            }
            Some(_) => {}
            None => {
                self.methods.insert(name.to_owned(), ty);
            }
        }
    }

    fn element(&self, name: &str) -> Option<&(Declared, String)> {
        self.elements.get(name)?.as_ref()
    }

    fn name(&self, name: &str) -> Option<&(Declared, String)> {
        self.names.get(name)?.as_ref()
    }
}

/// Rewrite every C# integral operator in `body` into its typed form.
pub fn normalize(body: &mut [Statement]) {
    let mut env = Env::default();
    for statement in body.iter() {
        collect_locals(statement, &mut env);
    }
    normalize_block(body, &env);
}

fn normalize_block(body: &mut [Statement], env: &Env) {
    for statement in body.iter_mut() {
        normalize_statement(statement, env);
    }
}

fn normalize_statement(statement: &mut Statement, env: &Env) {
    match &mut statement.kind {
        StmtKind::FunctionDecl { params, body, return_type, .. } => {
            let mut inner = env.clone();
            inner.returns = return_type.as_deref().and_then(declared).map(|d| d.ty);
            for param in params.iter() {
                inner.declare(&param.name, param.type_hint.as_deref());
            }
            for nested in body.iter() {
                collect_locals(nested, &mut inner);
            }
            normalize_block(body, &inner);
            return;
        }
        StmtKind::ClassDecl { members, .. } => {
            let mut inner = env.clone();
            for member in members.iter() {
                match member {
                    ClassMember::Field { name, type_hint, .. }
                    | ClassMember::Property { name, type_hint, .. } => {
                        inner.declare(name, type_hint.as_deref());
                    }
                    ClassMember::Const { name, type_hint, .. } => {
                        inner.declare(name, type_hint.as_deref());
                    }
                    ClassMember::Method(method) => {
                        if let StmtKind::FunctionDecl { name, return_type, .. } = &method.kind {
                            inner.declare_method(name, return_type.as_deref());
                        }
                    }
                    _ => {}
                }
            }
            for member in members.iter_mut() {
                match member {
                    ClassMember::Method(method) => normalize_statement(method, &inner),
                    ClassMember::Constructor { params, body, .. } => {
                        let mut ctor = inner.clone();
                        ctor.returns = None;
                        for param in params.iter() {
                            ctor.declare(&param.name, param.type_hint.as_deref());
                        }
                        for nested in body.iter() {
                            collect_locals(nested, &mut ctor);
                        }
                        normalize_block(body, &ctor);
                    }
                    ClassMember::Property { getter, .. } => {
                        if let Some(getter) = getter {
                            normalize_block(getter, &inner);
                        }
                    }
                    ClassMember::NestedType(nested) => normalize_statement(nested, &inner),
                    _ => {}
                }
            }
            return;
        }
        StmtKind::CompoundAssign { target, op, value } => {
            if let Some(assign) = typed_compound(target, *op, value, env) {
                statement.kind = assign;
                normalize_exprs(statement, env);
                return;
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for declaration in declarations.iter_mut() {
                typed_declaration(declaration, env);
            }
        }
        StmtKind::Expr(expr) => {
            if let Some(step) = typed_step(expr, env) {
                *expr = step;
            }
        }
        StmtKind::Return(Some(_)) => {
            normalize_exprs(statement, env);
            if let (Some(ty @ (IntTy::I64 | IntTy::U64)), StmtKind::Return(Some(value))) =
                (env.returns, &mut statement.kind)
            {
                *value = convert_value(std::mem::replace(value, Expression::new(ExprKind::Lit(Literal::Int(0)))), ty, env);
            }
            return;
        }
        StmtKind::For { init, update, body, .. } => {
            if let Some(init) = init {
                normalize_statement(init, env);
            }
            if let Some(update) = update {
                if let Some(step) = typed_step(update, env) {
                    *update = step;
                }
            }
            normalize_block(body, env);
        }
        StmtKind::Block(body)
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. }
        | StmtKind::ForIn { body, .. }
        | StmtKind::With { body, .. }
        | StmtKind::Using { body, .. }
        | StmtKind::Lock { body, .. }
        | StmtKind::NamespaceDecl { body, .. } => normalize_block(body, env),
        StmtKind::If { then_body, elifs, else_body, .. } => {
            normalize_block(then_body, env);
            for (_, branch) in elifs.iter_mut() {
                normalize_block(branch, env);
            }
            if let Some(branch) = else_body {
                normalize_block(branch, env);
            }
        }
        StmtKind::Switch { cases, default, .. } => {
            for case in cases.iter_mut() {
                normalize_block(&mut case.body, env);
            }
            if let Some(branch) = default {
                normalize_block(branch, env);
            }
        }
        StmtKind::Try { body, catches, else_body, finally } => {
            normalize_block(body, env);
            for catch in catches.iter_mut() {
                normalize_block(&mut catch.body, env);
            }
            if let Some(branch) = else_body {
                normalize_block(branch, env);
            }
            if let Some(branch) = finally {
                normalize_block(branch, env);
            }
        }
        StmtKind::Labeled { body, .. } => normalize_statement(body, env),
        _ => {}
    }
    normalize_exprs(statement, env);
}

/// Type the operators in this statement's own expressions. Nested statements
/// were handled by `normalize_statement`, and the rewrite is idempotent, so
/// the visitor crossing into them again changes nothing.
fn normalize_exprs(statement: &mut Statement, env: &Env) {
    statement.walk_exprs_mut(&mut |expr| {
        if let ExprKind::Lambda { body: LambdaBody::Block(body), params, .. } = &mut expr.kind {
            let mut inner = env.clone();
            for param in params.iter() {
                inner.declare(&param.name, param.type_hint.as_deref());
            }
            for nested in body.iter() {
                collect_locals(nested, &mut inner);
            }
            for nested in body.iter_mut() {
                normalize_statement_steps(nested, &inner);
            }
            return;
        }
        if let Some(typed) = typed_expression(expr, env) {
            *expr = typed;
        }
    });
}

/// Statement-position steps and compound assignments inside a lambda block,
/// whose expressions the enclosing visitor already reaches.
fn normalize_statement_steps(statement: &mut Statement, env: &Env) {
    match &mut statement.kind {
        StmtKind::CompoundAssign { target, op, value } => {
            if let Some(assign) = typed_compound(target, *op, value, env) {
                statement.kind = assign;
            }
        }
        StmtKind::Expr(expr) => {
            if let Some(step) = typed_step(expr, env) {
                *expr = step;
            }
        }
        _ => {}
    }
}

/// Record the declared type of every local this statement introduces,
/// without entering nested function or class declarations.
fn collect_locals(statement: &Statement, env: &mut Env) {
    match &statement.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for declaration in declarations {
                if let vybe_ast::BindingPattern::Ident(name) = &declaration.pattern {
                    match (declaration.type_hint.as_deref(), &declaration.init) {
                        (None, Some(init)) => {
                            let inferred = type_of(init, env).map(IntTy::spelling);
                            env.declare(name, inferred);
                        }
                        (hint, _) => env.declare(name, hint),
                    }
                }
            }
        }
        StmtKind::ForIn { var, key, body, iter, .. } => {
            let element = match &iter.kind {
                ExprKind::Ident(name) => env.element(name).map(|(_, spelling)| spelling.clone()),
                _ => None,
            };
            env.declare(var, element.as_deref());
            if let Some(key) = key {
                env.declare(key, None);
            }
            body.iter().for_each(|nested| collect_locals(nested, env));
        }
        StmtKind::For { init, body, .. } => {
            if let Some(init) = init {
                collect_locals(init, env);
            }
            body.iter().for_each(|nested| collect_locals(nested, env));
        }
        StmtKind::Block(body)
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. }
        | StmtKind::With { body, .. }
        | StmtKind::Using { body, .. }
        | StmtKind::Lock { body, .. } => body.iter().for_each(|nested| collect_locals(nested, env)),
        StmtKind::If { then_body, elifs, else_body, .. } => {
            then_body.iter().for_each(|nested| collect_locals(nested, env));
            for (_, branch) in elifs {
                branch.iter().for_each(|nested| collect_locals(nested, env));
            }
            if let Some(branch) = else_body {
                branch.iter().for_each(|nested| collect_locals(nested, env));
            }
        }
        StmtKind::Switch { cases, default, .. } => {
            for case in cases {
                case.body.iter().for_each(|nested| collect_locals(nested, env));
            }
            if let Some(branch) = default {
                branch.iter().for_each(|nested| collect_locals(nested, env));
            }
        }
        StmtKind::Try { body, catches, else_body, finally } => {
            body.iter().for_each(|nested| collect_locals(nested, env));
            for catch in catches {
                if let Some(name) = &catch.var_name {
                    env.declare(name, None);
                }
                catch.body.iter().for_each(|nested| collect_locals(nested, env));
            }
            for branch in [else_body, finally].into_iter().flatten() {
                branch.iter().for_each(|nested| collect_locals(nested, env));
            }
        }
        StmtKind::Labeled { body, .. } => collect_locals(body, env),
        StmtKind::FunctionDecl { name, return_type, .. } => {
            env.declare_method(name, return_type.as_deref());
        }
        _ => {}
    }
}

/// The C# integral type of `expr`, or `None` when it is not integral or not
/// known from declarations.
fn type_of(expr: &Expression, env: &Env) -> Option<IntTy> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(value)) => {
            Some(if i32::try_from(*value).is_ok() { IntTy::I32 } else { IntTy::I64 })
        }
        ExprKind::Cast { type_name, .. } => declared(type_name).map(|d| d.ty),
        ExprKind::Ident(name) => env.name(name).map(|(d, _)| d.ty),
        ExprKind::Index { object, .. } => match &object.kind {
            ExprKind::Ident(name) => env.element(name).map(|(d, _)| d.ty),
            _ => None,
        },
        ExprKind::Member { object, field, .. }
            if matches!(&object.kind, ExprKind::This) =>
        {
            env.name(field).map(|(d, _)| d.ty)
        }
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(name) => env.methods.get(name).copied().flatten(),
            _ => None,
        },
        ExprKind::Binary { op: BinOp::Integer(operation, lane), .. } => {
            if operation.is_comparison() {
                None
            } else {
                Some(match lane {
                    BitLane::W64 => IntTy::I64,
                    _ => IntTy::I32,
                })
            }
        }
        ExprKind::Binary { op, left, right } => {
            let (left, right) = (type_of(left, env)?, type_of(right, env)?);
            match op {
                BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::Div | BinOp::IDiv | BinOp::Mod
                | BinOp::BitAnd | BinOp::BitOr | BinOp::BitXor => promote(left, right, None),
                BinOp::Shl | BinOp::Shr | BinOp::UShr => Some(left),
                _ => None,
            }
        }
        ExprKind::Ternary { then, else_, .. } => {
            if is_throw_expression(then) {
                return type_of(else_, env);
            }
            if is_throw_expression(else_) {
                return type_of(then, env);
            }
            let ty = type_of(then, env)?;
            (type_of(else_, env)? == ty).then_some(ty)
        }
        ExprKind::Unary { op: UnaryOp::Neg, expr } => match type_of(expr, env)? {
            IntTy::I32 => Some(IntTy::I32),
            IntTy::U32 | IntTy::I64 => Some(IntTy::I64),
            IntTy::U64 => None,
        },
        _ => None,
    }
}

/// A non-negative integer constant, which converts implicitly to an unsigned
/// type (ECMA-334 §10.2.11).
fn non_negative_constant(expr: &Expression) -> bool {
    matches!(&expr.kind, ExprKind::Lit(Literal::Int(value)) if *value >= 0)
}

/// ECMA-334 §12.4.7.3 binary numeric promotion over integral operands.
/// `constants` says which operands are non-negative constants.
fn promote(left: IntTy, right: IntTy, constants: Option<(bool, bool)>) -> Option<IntTy> {
    use IntTy::*;
    let (left_const, right_const) = constants.unwrap_or((false, false));
    Some(match (left, right) {
        (U64, U64) | (U64, U32) | (U32, U64) => U64,
        (U64, _) if right_const => U64,
        (_, U64) if left_const => U64,
        (U64, _) | (_, U64) => return None,
        (I64, _) | (_, I64) => I64,
        (U32, U32) => U32,
        (U32, I32) if right_const => U32,
        (I32, U32) if left_const => U32,
        (U32, I32) | (I32, U32) => I64,
        (I32, I32) => I32,
    })
}

fn int_op(op: &BinOp, ty: IntTy, right: &Expression) -> Option<IntOp> {
    let signed = ty.signed();
    let divisor_ok = matches!(&right.kind,
        ExprKind::Lit(Literal::Int(value)) if *value != 0 && *value != -1);
    Some(match op {
        BinOp::Add => IntOp::Add,
        BinOp::Sub => IntOp::Sub,
        BinOp::Mul => IntOp::Mul,
        BinOp::BitAnd => IntOp::And,
        BinOp::BitOr => IntOp::Or,
        BinOp::BitXor => IntOp::Xor,
        BinOp::Shl => IntOp::Shl,
        BinOp::Shr => if signed { IntOp::ShrS } else { IntOp::ShrU },
        BinOp::UShr => IntOp::ShrU,
        BinOp::Div | BinOp::IDiv if divisor_ok => if signed { IntOp::DivS } else { IntOp::DivU },
        BinOp::Mod if divisor_ok => if signed { IntOp::RemS } else { IntOp::RemU },
        BinOp::Eq | BinOp::StrictEq => IntOp::Eq,
        BinOp::NotEq | BinOp::StrictNotEq => IntOp::Ne,
        BinOp::Lt => if signed { IntOp::LtS } else { IntOp::LtU },
        BinOp::LtEq => if signed { IntOp::LeS } else { IntOp::LeU },
        BinOp::Gt => if signed { IntOp::GtS } else { IntOp::GtU },
        BinOp::GtEq => if signed { IntOp::GeS } else { IntOp::GeU },
        _ => return None,
    })
}

/// `expr` converted to `ty`: unchanged when it already has that type.
fn convert(expr: Expression, ty: IntTy, env: &Env) -> Expression {
    if type_of(&expr, env) == Some(ty) && !matches!(&expr.kind, ExprKind::Lit(_)) {
        return expr;
    }
    if ty == IntTy::I32 && matches!(&expr.kind, ExprKind::Lit(Literal::Int(value)) if i32::try_from(*value).is_ok()) {
        return expr;
    }
    Expression::with_span(
        ExprKind::Cast { expr: Box::new(expr.clone()), type_name: ty.spelling().to_owned() },
        expr.span,
    )
}

/// A value converted to `ty`, the conversion distributed over the arms of a
/// conditional so each arm states its own type.
fn convert_value(expr: Expression, ty: IntTy, env: &Env) -> Expression {
    match expr.kind {
        ExprKind::Ternary { cond, then, else_ } => Expression::with_span(
            ExprKind::Ternary {
                cond,
                then: Box::new(convert_value(*then, ty, env)),
                else_: Box::new(convert_value(*else_, ty, env)),
            },
            expr.span,
        ),
        kind => convert(Expression::with_span(kind, expr.span), ty, env),
    }
}

/// The typed form of an integral binary operator, or `None`.
fn typed_binary(op: &BinOp, left: &Expression, right: &Expression, env: &Env) -> Option<Expression> {
    let (left_ty, right_ty) = (type_of(left, env)?, type_of(right, env)?);
    let shift = matches!(op, BinOp::Shl | BinOp::Shr | BinOp::UShr);
    let ty = if shift {
        left_ty
    } else {
        promote(left_ty, right_ty, Some((non_negative_constant(left), non_negative_constant(right))))?
    };
    if matches!(op, BinOp::Div | BinOp::IDiv | BinOp::Mod) && int_op(op, ty, right).is_none() {
        return guarded_division(op, ty, left, right, env);
    }
    let operation = int_op(op, ty, right)?;
    let node = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::Integer(operation, ty.lane()),
            left: Box::new(convert(left.clone(), ty, env)),
            right: Box::new(convert(right.clone(), ty, env)),
        },
        left.span,
    );
    Some(if !operation.is_comparison() && !ty.signed() {
        Expression::with_span(
            ExprKind::Cast { expr: Box::new(node), type_name: ty.spelling().to_owned() },
            left.span,
        )
    } else {
        node
    })
}

/// An operand that may be evaluated more than once.
fn pure(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Ident(_) | ExprKind::Lit(_) => true,
        ExprKind::Cast { expr, .. } => pure(expr),
        ExprKind::Member { object, .. } => matches!(&object.kind, ExprKind::This),
        _ => false,
    }
}

/// `throw new <class>(message)` in expression position.
fn throw_expression(class: &str, message: &str, span: vybe_ast::Span) -> Expression {
    let throw = Statement::with_span(
        StmtKind::Throw {
            expr: Some(Expression::with_span(
                ExprKind::New {
                    class: Box::new(Expression::ident(class)),
                    args: vec![vybe_ast::Argument::positional(Expression::with_span(
                        ExprKind::Lit(Literal::Str(message.into())),
                        span,
                    ))],
                },
                span,
            )),
            cause: None,
        },
        span,
    );
    let lambda = Expression::with_span(
        ExprKind::Lambda {
            params: vec![],
            body: LambdaBody::Block(vec![throw]),
            is_async: false,
            captures: Vec::new(),
        },
        span,
    );
    Expression::with_span(ExprKind::Call { callee: Box::new(lambda), args: vec![], optional: false }, span)
}

fn is_throw_expression(expr: &Expression) -> bool {
    matches!(&expr.kind, ExprKind::Call { callee, args, .. } if args.is_empty() && matches!(
        &callee.kind,
        ExprKind::Lambda { params, body: LambdaBody::Block(body), .. }
            if params.is_empty() && matches!(body.as_slice(), [Statement { kind: StmtKind::Throw { .. }, .. }])
    ))
}

/// Integral `/` and `%` by a divisor that is not a safe constant
/// (ECMA-334 §12.10.3, §12.10.4): a zero divisor throws
/// `DivideByZeroException`, and the signed `MinValue / -1` and `MinValue % -1`
/// throw `OverflowException`. Every other quotient is the wasm instruction's.
fn guarded_division(op: &BinOp, ty: IntTy, left: &Expression, right: &Expression, env: &Env) -> Option<Expression> {
    if !pure(right) || (ty.signed() && !pure(left)) {
        return None;
    }
    let span = left.span;
    let signed = ty.signed();
    let operation = match (op, signed) {
        (BinOp::Mod, true) => IntOp::RemS,
        (BinOp::Mod, false) => IntOp::RemU,
        (_, true) => IntOp::DivS,
        (_, false) => IntOp::DivU,
    };
    let (left, right) = (convert(left.clone(), ty, env), convert(right.clone(), ty, env));
    let constant = |value: i64| convert(Expression::with_span(ExprKind::Lit(Literal::Int(value)), span), ty, env);
    let test = |a: &Expression, value: i64| Box::new(Expression::with_span(
        ExprKind::Binary {
            op: BinOp::Integer(IntOp::Eq, ty.lane()),
            left: Box::new(a.clone()),
            right: Box::new(constant(value)),
        },
        span,
    ));
    let ternary = |cond, then: Expression, else_: Expression| Expression::with_span(
        ExprKind::Ternary { cond, then: Box::new(then), else_: Box::new(else_) },
        span,
    );
    let mut quotient = Expression::with_span(
        ExprKind::Binary {
            op: BinOp::Integer(operation, ty.lane()),
            left: Box::new(left.clone()),
            right: Box::new(right.clone()),
        },
        span,
    );
    if !signed {
        quotient = Expression::with_span(
            ExprKind::Cast { expr: Box::new(quotient), type_name: ty.spelling().to_owned() },
            span,
        );
    }
    if signed {
        let minimum = if ty == IntTy::I64 { i64::MIN } else { i64::from(i32::MIN) };
        quotient = ternary(
            test(&right, -1),
            ternary(
                test(&left, minimum),
                throw_expression("OverflowException", "Arithmetic operation resulted in an overflow.", span),
                quotient.clone(),
            ),
            quotient,
        );
    }
    Some(ternary(
        test(&right, 0),
        throw_expression("DivideByZeroException", "Attempted to divide by zero.", span),
        quotient,
    ))
}

fn typed_expression(expr: &Expression, env: &Env) -> Option<Expression> {
    match &expr.kind {
        ExprKind::Binary { op, left, right } if !matches!(op, BinOp::Integer(..)) => {
            let typed = typed_binary(op, left, right, env)?;
            Some(if is_comparison(&typed) { boolean(typed) } else { typed })
        }
        ExprKind::Member { object, field, .. } => {
            let owner = match &object.kind {
                ExprKind::Ident(owner) => owner.as_str(),
                ExprKind::Member { object: system, field: owner, .. }
                    if matches!(&system.kind, ExprKind::Ident(root) if root == "System") =>
                {
                    owner.as_str()
                }
                _ => return None,
            };
            wide_limit(owner, field).map(|(bits, ty)| {
                Expression::with_span(
                    ExprKind::Cast {
                        expr: Box::new(Expression::with_span(ExprKind::Lit(Literal::Int(bits)), expr.span)),
                        type_name: ty.spelling().to_owned(),
                    },
                    expr.span,
                )
            })
        }
        // `c ? true : false` around a comparison this pass already
        // materialized: the condition is the comparison itself.
        ExprKind::Ternary { cond, then, else_ }
            if is_bool(then, true) && is_bool(else_, false) =>
        {
            let ExprKind::Ternary { cond: inner, then: t, else_: e } = &cond.kind else {
                return None;
            };
            (is_comparison(inner) && is_bool(t, true) && is_bool(e, false)).then(|| {
                Expression::with_span(
                    ExprKind::Ternary { cond: inner.clone(), then: then.clone(), else_: else_.clone() },
                    expr.span,
                )
            })
        }
        ExprKind::Unary { op: op @ (UnaryOp::PreInc | UnaryOp::PreDec), expr: target } => {
            let value = step_value(target, matches!(op, UnaryOp::PreInc), env)?;
            Some(Expression::with_span(
                ExprKind::Assign { target: target.clone(), value: Box::new(value) },
                expr.span,
            ))
        }
        _ => None,
    }
}

/// `long`/`ulong` `MinValue` and `MaxValue` as exact typed constants: their
/// values are not representable as a double.
fn wide_limit(owner: &str, field: &str) -> Option<(i64, IntTy)> {
    match (owner, field) {
        ("long" | "Int64", "MaxValue") => Some((i64::MAX, IntTy::I64)),
        ("long" | "Int64", "MinValue") => Some((i64::MIN, IntTy::I64)),
        ("ulong" | "UInt64", "MaxValue") => Some((u64::MAX as i64, IntTy::U64)),
        ("ulong" | "UInt64", "MinValue") => Some((0, IntTy::U64)),
        _ => None,
    }
}

fn is_comparison(expr: &Expression) -> bool {
    matches!(&expr.kind, ExprKind::Binary { op: BinOp::Integer(operation, _), .. } if operation.is_comparison())
}

fn is_bool(expr: &Expression, value: bool) -> bool {
    matches!(&expr.kind, ExprKind::Lit(Literal::Bool(b)) if *b == value)
}

/// An integer comparison's i32 predicate as a C# `bool` value.
fn boolean(comparison: Expression) -> Expression {
    let span = comparison.span;
    Expression::with_span(
        ExprKind::Ternary {
            cond: Box::new(comparison),
            then: Box::new(Expression::with_span(ExprKind::Lit(Literal::Bool(true)), span)),
            else_: Box::new(Expression::with_span(ExprKind::Lit(Literal::Bool(false)), span)),
        },
        span,
    )
}

fn is_floating(spelling: &str) -> bool {
    let spelling = spelling.trim();
    matches!(spelling.strip_prefix("System.").unwrap_or(spelling), "double" | "float" | "Double" | "Single")
}

/// A local declaration's typed form: a 64-bit integral local states the
/// `Checked` fact its storage relies on (the C# compiler enforces the declared
/// type, and a local is never `volatile`), a `var` takes the type of its
/// initializer, and an integral initializer of a floating local is converted
/// (ECMA-334 §10.2.3).
fn typed_declaration(declaration: &mut vybe_ast::VarDeclarator, env: &Env) {
    let wide = |ty: IntTy| matches!(ty, IntTy::I64 | IntTy::U64);
    let checked = |ty: IntTy| {
        let mut hint = vybe_ast::TypeHint::checked(ty.spelling());
        hint.access = Some(vybe_ast::StorageAccess::Ordinary);
        hint
    };
    let init_ty = declaration.init.as_ref().and_then(|init| type_of(init, env));
    match declaration.type_hint.as_ref() {
        None => {
            if let Some(ty) = init_ty.filter(|ty| wide(*ty)) {
                declaration.type_hint = Some(checked(ty));
            }
        }
        Some(hint) if hint.binding != vybe_ast::TypeBinding::Checked => {
            if let Some(ty) = declared(hint).filter(|d| !d.narrow).map(|d| d.ty).filter(|ty| wide(*ty)) {
                declaration.type_hint = Some(checked(ty));
                if let Some(init) = declaration.init.take() {
                    declaration.init = Some(convert(init, ty, env));
                }
            } else if is_floating(hint) && init_ty.is_some() {
                if let Some(init) = declaration.init.take() {
                    let span = init.span;
                    declaration.init = Some(Expression::with_span(
                        ExprKind::Cast { expr: Box::new(init), type_name: "double".to_owned() },
                        span,
                    ));
                }
            }
        }
        Some(_) => {}
    }
}

/// `target ± 1` in `target`'s declared type, for a plain local or field.
fn step_value(target: &Expression, increment: bool, env: &Env) -> Option<Expression> {
    let ExprKind::Ident(name) = &target.kind else {
        return None;
    };
    let (declared, spelling) = env.name(name)?.clone();
    let one = Expression::with_span(ExprKind::Lit(Literal::Int(1)), target.span);
    let op = if increment { BinOp::Add } else { BinOp::Sub };
    let sum = typed_binary(&op, target, &one, env)?;
    Some(if declared.narrow {
        Expression::with_span(ExprKind::Cast { expr: Box::new(sum), type_name: spelling }, target.span)
    } else {
        sum
    })
}

/// A statement-position `x++` / `x--` / `++x` / `--x` as `x = x ± 1`; its
/// value is discarded, so the post-increment's old value is never needed.
fn typed_step(expr: &Expression, env: &Env) -> Option<Expression> {
    let ExprKind::Unary { op, expr: target } = &expr.kind else {
        return None;
    };
    let increment = match op {
        UnaryOp::PostInc | UnaryOp::PreInc => true,
        UnaryOp::PostDec | UnaryOp::PreDec => false,
        _ => return None,
    };
    let value = step_value(target, increment, env)?;
    Some(Expression::with_span(
        ExprKind::Assign { target: target.clone(), value: Box::new(value) },
        expr.span,
    ))
}

/// `x op= v` as `x = (T)(x op v)` for a plain integral local or field
/// (ECMA-334 §12.21.4).
fn typed_compound(target: &Expression, op: CompoundOp, value: &Expression, env: &Env) -> Option<StmtKind> {
    let ExprKind::Ident(name) = &target.kind else {
        return None;
    };
    let (declared, spelling) = env.name(name)?.clone();
    let op = match op {
        CompoundOp::Add => BinOp::Add,
        CompoundOp::Sub => BinOp::Sub,
        CompoundOp::Mul => BinOp::Mul,
        CompoundOp::Div => BinOp::Div,
        CompoundOp::IDiv => BinOp::IDiv,
        CompoundOp::Mod => BinOp::Mod,
        CompoundOp::BitAnd => BinOp::BitAnd,
        CompoundOp::BitOr => BinOp::BitOr,
        CompoundOp::BitXor => BinOp::BitXor,
        CompoundOp::Shl => BinOp::Shl,
        CompoundOp::Shr => BinOp::Shr,
        CompoundOp::UShr => BinOp::UShr,
        _ => return None,
    };
    let sum = typed_binary(&op, target, value, env)?;
    let result = type_of(&sum, env);
    let value = if declared.narrow || result != Some(declared.ty) {
        Expression::with_span(
            ExprKind::Cast { expr: Box::new(sum), type_name: spelling },
            target.span,
        )
    } else {
        sum
    };
    Some(StmtKind::Assign { targets: vec![target.clone()], value, by_ref: false })
}
