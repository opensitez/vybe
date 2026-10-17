//! Selected closure analysis fusion. The reference helpers retain nested closure scope
//! filtering; this visitor shares only the enclosing body's discovery traversal.
//! No AST fields, memo table, source scan, or per-node allocation.
use std::collections::HashSet;
use vybe_ast::*;
use super as old;

pub(super) fn analyze(body: &[Statement], captured: &mut HashSet<String>) -> bool {
    let mut state = State { captured, receiver: false };
    state.body(body, true, true);
    state.receiver
}

struct State<'a> {
    captured: &'a mut HashSet<String>,
    receiver: bool,
}
impl State<'_> {
    fn body(&mut self, body: &[Statement], capture: bool, receiver: bool) {
        for stmt in body { self.stmt(stmt, capture, receiver); }
    }
    fn stmt(&mut self, stmt: &Statement, capture: bool, receiver: bool) {
        let receiver = receiver && !self.receiver;
        if !capture {
            if receiver { self.receiver = old::stmt_has_closure_with_this(stmt); }
            return;
        }
        if !receiver {
            old::collect_closure_captured_in_stmt(stmt, self.captured);
            return;
        }
        match &stmt.kind {
            // The receiver analysis stops at ordinary functions/classes.
            StmtKind::FunctionDecl { .. } | StmtKind::ClassDecl { .. } => {
                old::collect_closure_captured_in_stmt(stmt, self.captured);
            }
            StmtKind::Expr(expr) | StmtKind::Return(Some(expr))
            | StmtKind::Throw { expr: Some(expr), .. } => self.expr(expr, true, true),
            StmtKind::VarDecl { declarations, .. } => {
                for d in declarations {
                    if let Some(expr) = &d.init { self.expr(expr, true, true); }
                }
            }
            StmtKind::Assign { targets, value, .. } => {
                for target in targets { self.expr(target, true, false); }
                self.expr(value, true, true);
            }
            StmtKind::CompoundAssign { target, value, .. } => {
                self.expr(target, true, false);
                self.expr(value, true, true);
            }
            StmtKind::Block(body) => self.body(body, true, true),
            StmtKind::If { cond, then_body, elifs, else_body } => {
                self.expr(cond, true, true);
                self.body(then_body, true, true);
                for (cond, body) in elifs {
                    self.expr(cond, true, true);
                    self.body(body, true, true);
                }
                if let Some(body) = else_body { self.body(body, true, true); }
            }
            StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
                self.expr(cond, true, true);
                self.body(body, true, true);
            }
            StmtKind::For { init, cond, update, body, .. } => {
                if let Some(stmt) = init { self.stmt(stmt, true, true); }
                if let Some(expr) = cond { self.expr(expr, true, true); }
                if let Some(expr) = update { self.expr(expr, true, true); }
                self.body(body, true, true);
            }
            StmtKind::ForIn { iter, body, .. } => {
                self.expr(iter, true, true);
                self.body(body, true, true);
            }
            StmtKind::Try { body, catches, else_body, finally } => {
                self.body(body, true, true);
                for c in catches { self.body(&c.body, true, true); }
                // The original receiver analysis intentionally omits else_body.
                if let Some(body) = else_body { self.body(body, true, false); }
                if let Some(body) = finally { self.body(body, true, true); }
            }
            StmtKind::Switch { expr, cases, default } => {
                self.expr(expr, true, true);
                for c in cases { self.body(&c.body, true, true); }
                if let Some(body) = default { self.body(body, true, true); }
            }
            StmtKind::Labeled { body, .. } => self.stmt(body, true, true),
            _ => {}
        }
    }
    fn expr(&mut self, expr: &Expression, capture: bool, receiver: bool) {
        let receiver = receiver && !self.receiver;
        if !capture {
            if receiver { self.receiver = old::expr_has_closure_with_this(expr); }
            return;
        }
        if !receiver {
            old::collect_closure_captured_in_expr(expr, self.captured);
            return;
        }
        match &expr.kind {
            // Retain each original nested analysis and its different scope
            // boundaries. Stage one does not claim to fuse the inner scans.
            ExprKind::Lambda { .. } => {
                old::collect_closure_captured_in_expr(expr, self.captured);
                self.receiver = old::expr_has_closure_with_this(expr);
            }
            ExprKind::FunctionExpr(_) | ExprKind::ClassExpr { .. } | ExprKind::RefOf(_) => {
                old::collect_closure_captured_in_expr(expr, self.captured);
            }
            ExprKind::Unary { expr, .. } | ExprKind::Await(expr) | ExprKind::Spread(expr) => {
                self.expr(expr, true, true);
            }
            ExprKind::Cast { expr, .. } | ExprKind::IsType { expr, .. }
            | ExprKind::RefLoad(expr) | ExprKind::Yield(Some(expr)) => {
                self.expr(expr, true, false);
            }
            ExprKind::Binary { left, right, .. }
            | ExprKind::Assign { target: left, value: right } => {
                self.expr(left, true, true);
                self.expr(right, true, true);
            }
            ExprKind::Ternary { cond, then, else_ } => {
                self.expr(cond, true, true);
                self.expr(then, true, true);
                self.expr(else_, true, true);
            }
            ExprKind::Call { callee, args, .. } => {
                self.expr(callee, true, true);
                for arg in args { self.expr(&arg.value, true, true); }
            }
            ExprKind::Member { object, .. } => self.expr(object, true, true),
            ExprKind::Index { object, index, .. } => {
                self.expr(object, true, true);
                // Preserve the original receiver visitor's object-only policy.
                self.expr(index, true, false);
            }
            ExprKind::Array(elems) => {
                for elem in elems { self.expr(&elem.value, true, true); }
            }
            ExprKind::Object(props) => {
                for prop in props {
                    match prop {
                        ObjectProperty::KeyValue { value, .. } => self.expr(value, true, true),
                        ObjectProperty::Computed { key, value } => {
                            self.expr(key, true, false);
                            self.expr(value, true, false);
                        }
                        ObjectProperty::Method { value, .. }
                        | ObjectProperty::Accessor { value, .. } => {
                            if matches!(value.kind, StmtKind::FunctionDecl { .. }) {
                                old::collect_closure_captured_in_stmt(value, self.captured);
                            }
                        }
                        _ => {}
                    }
                }
            }
            ExprKind::Sequence(exprs) => {
                for expr in exprs { self.expr(expr, true, true); }
            }
            ExprKind::Tuple(exprs) => {
                for expr in exprs { self.expr(expr, true, false); }
            }
            ExprKind::Interpolation(parts) => {
                for part in parts {
                    if let InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) = part {
                        self.expr(expr, true, false);
                    }
                }
            }
            ExprKind::Async(op) => {
                for child in op.children() { self.expr(child, true, true); }
            }
            ExprKind::Chan(op) => {
                for child in op.children() { self.expr(child, true, true); }
            }
            ExprKind::New { class, args } => {
                self.expr(class, true, true);
                for arg in args { self.expr(&arg.value, true, true); }
            }
            _ => {}
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn compare(body: &[Statement]) {
        let mut expected=HashSet::from(["<existing capture>".to_owned()]);
        old::collect_closure_captured_idents(body,&mut expected);
        let receiver=old::closures_in_body_reference_this(body);
        let mut actual=HashSet::from(["<existing capture>".to_owned()]);
        assert_eq!(super::analyze(body,&mut actual),receiver);
        assert_eq!(actual,expected);
    }

    #[test]
    fn capture_and_receiver_keep_distinct_common_ast_barriers() {
        let arrow=Expression::new(ExprKind::Lambda {
            params:vec![],body:LambdaBody::Expr(Box::new(Expression::new(
                ExprKind::Member {object:Box::new(Expression::ident("this")),
                    field:"x".into(),null_safe:false}))),
            is_async:false,captures:vec![],
        });
        let cases=[
            (ExprKind::Cast{expr:Box::new(arrow.clone()),type_name:"T".into()},false),
            (ExprKind::IsType{expr:Box::new(arrow.clone()),type_name:"T".into()},false),
            (ExprKind::RefLoad(Box::new(arrow.clone())),false),
            (ExprKind::RefOf(Box::new(PlaceExpr::Deref(Box::new(arrow.clone())))),false),
            (ExprKind::Yield(Some(Box::new(arrow.clone()))),false),
            (ExprKind::TypeOf(Box::new(arrow.clone())),false),
            (ExprKind::Delete(Box::new(arrow.clone())),false),
            (ExprKind::NullCoalesce{left:Box::new(arrow.clone()),right:Box::new(Expression::int(0))},false),
            (ExprKind::Tuple(vec![arrow.clone()]),false),
            (ExprKind::Interpolation(vec![InterpolPart::Expr(arrow.clone())]),false),
            (ExprKind::Object(vec![ObjectProperty::Computed{key:arrow.clone(),value:arrow.clone()}]),false),
            (ExprKind::Index{object:Box::new(Expression::ident("obj")),index:Box::new(arrow.clone()),null_safe:false},false),
            (ExprKind::Index{object:Box::new(arrow.clone()),index:Box::new(Expression::int(0)),null_safe:false},true),
            (ExprKind::Await(Box::new(arrow.clone())),true),
            (ExprKind::Spread(Box::new(arrow.clone())),true),
        ];
        for (kind,receiver) in cases {
            let body=vec![Statement::new(StmtKind::Expr(Expression::new(kind)))];
            assert_eq!(old::closures_in_body_reference_this(&body),receiver);
            compare(&body);
        }
        let expr=Statement::new(StmtKind::Expr(arrow));
        for (else_body,finally,receiver) in [(Some(vec![expr.clone()]),None,false),(None,Some(vec![expr]),true)] {
            let body=vec![Statement::new(StmtKind::Try{body:vec![],catches:vec![],else_body,finally})];
            assert_eq!(old::closures_in_body_reference_this(&body),receiver);
            compare(&body);
        }
        compare(&[]);
    }
}
