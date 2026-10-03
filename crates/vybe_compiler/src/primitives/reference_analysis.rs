//! Resolve module address-taking before emitting any function bodies.
use std::collections::{HashMap, HashSet};
use vybe_ast::*;

#[derive(Default)]
struct Frame {
    locals: HashSet<String>,
    redirects: HashMap<String, ScopeDeclKind>,
    closed: bool,
    function: bool,
}

pub(super) fn module_address_taken(
    body: &[Statement],
    fold: Option<CaseAlphabet>,
) -> HashSet<String> {
    let mut scan = Scan {
        frames: vec![Frame {
            function: true,
            ..Frame::default()
        }],
        globals: HashSet::new(),
        fold,
    };
    scan.body(body);
    scan.globals
}

struct Scan {
    frames: Vec<Frame>,
    globals: HashSet<String>,
    fold: Option<CaseAlphabet>,
}

impl Scan {
    fn key(&self, name: &str) -> String {
        match self.fold {
            None => name.to_owned(),
            Some(CaseAlphabet::Ascii) => name.to_ascii_lowercase(),
            Some(CaseAlphabet::Unicode) => name.to_lowercase(),
        }
    }

    fn address(&mut self, name: &str) {
        let name = self.key(name);
        for frame in self.frames.iter().skip(1).rev() {
            match frame.redirects.get(&name) {
                Some(ScopeDeclKind::Global) => break,
                Some(ScopeDeclKind::Nonlocal) => continue,
                _ => {}
            }
            if frame.locals.contains(&name) || frame.closed {
                return;
            }
        }
        self.globals.insert(name.to_owned());
    }

    fn bind(&mut self, name: &str, function_scoped: bool) {
        let name = self.key(name);
        let frame = if function_scoped {
            self.frames.iter_mut().rev().find(|f| f.function).unwrap()
        } else {
            self.frames.last_mut().unwrap()
        };
        frame.locals.insert(name);
    }

    fn pattern(&mut self, pattern: &BindingPattern, function_scoped: bool) {
        match pattern {
            BindingPattern::Ident(name) => self.bind(name, function_scoped),
            BindingPattern::Object(props) => {
                for p in props {
                    if let Some(value) = &p.value {
                        self.pattern(value, function_scoped);
                    } else {
                        self.bind(&p.key, function_scoped);
                    }
                }
            }
            BindingPattern::Array(items) => {
                for item in items {
                    match item {
                        ArrayPatternElem::Pattern(p, _) => self.pattern(p, function_scoped),
                        ArrayPatternElem::Rest(name) => self.bind(name, function_scoped),
                        ArrayPatternElem::Hole => {}
                    }
                }
            }
        }
    }

    fn body(&mut self, body: &[Statement]) {
        // Lexical bindings shadow throughout their block, including before
        // initialization. Ordinary declarations become visible in order.
        for stmt in body {
            if let StmtKind::VarDecl {
                declarations,
                kind: VarDeclKind::Let | VarDeclKind::Const,
            } = &stmt.kind
            {
                for decl in declarations {
                    self.pattern(&decl.pattern, false);
                }
            }
        }
        for stmt in body {
            self.stmt(stmt);
        }
    }

    fn block(&mut self, body: &[Statement]) {
        self.frames.push(Frame::default());
        self.body(body);
        self.frames.pop();
    }

    fn function(&mut self, params: &[Param], body: &[Statement]) {
        self.frames.push(Frame {
            function: true,
            ..Frame::default()
        });
        for param in params {
            self.bind(&param.name, false);
            if let Some(default) = &param.default {
                self.expr(default);
            }
        }
        self.function_bindings(body);
        self.body(body);
        self.frames.pop();
    }

    fn function_bindings(&mut self, body: &[Statement]) {
        for stmt in body {
            match &stmt.kind {
                StmtKind::VarDecl {
                    declarations,
                    kind: VarDeclKind::FunctionScoped,
                } => {
                    for decl in declarations {
                        self.pattern(&decl.pattern, true);
                    }
                }
                StmtKind::Block(body)
                | StmtKind::While { body, .. }
                | StmtKind::DoWhile { body, .. }
                | StmtKind::ForIn { body, .. } => self.function_bindings(body),
                StmtKind::For { init, body, .. } => {
                    if let Some(init) = init {
                        self.function_bindings(std::slice::from_ref(init.as_ref()));
                    }
                    self.function_bindings(body);
                }
                StmtKind::If {
                    then_body,
                    elifs,
                    else_body,
                    ..
                } => {
                    self.function_bindings(then_body);
                    for (_, body) in elifs {
                        self.function_bindings(body);
                    }
                    if let Some(body) = else_body {
                        self.function_bindings(body);
                    }
                }
                StmtKind::Try {
                    body,
                    catches,
                    else_body,
                    finally,
                } => {
                    self.function_bindings(body);
                    for catch in catches {
                        self.function_bindings(&catch.body);
                    }
                    for body in [else_body, finally].into_iter().flatten() {
                        self.function_bindings(body);
                    }
                }
                StmtKind::Switch { cases, default, .. } => {
                    for case in cases {
                        self.function_bindings(&case.body);
                    }
                    if let Some(body) = default {
                        self.function_bindings(body);
                    }
                }
                _ => {}
            }
        }
    }

    fn stmt(&mut self, stmt: &Statement) {
        match &stmt.kind {
            StmtKind::FunctionDecl {
                name, params, body, ..
            } => {
                self.bind(name, false);
                self.function(params, body);
            }
            StmtKind::ScopeDecl { kind, names } => {
                let names: Vec<_> = names.iter().map(|name| self.key(name)).collect();
                let frame = self.frames.iter_mut().rev().find(|f| f.function).unwrap();
                if *kind == ScopeDeclKind::Closed {
                    frame.closed = true;
                }
                for name in names {
                    frame.redirects.insert(
                        name.clone(),
                        if *kind == ScopeDeclKind::Closed {
                            ScopeDeclKind::Global
                        } else {
                            *kind
                        },
                    );
                }
            }
            StmtKind::VarDecl { declarations, kind } => {
                for d in declarations {
                    self.pattern(&d.pattern, *kind == VarDeclKind::FunctionScoped);
                    if let Some(init) = &d.init {
                        self.expr(init);
                    }
                    if let Some(bounds) = &d.array_bounds {
                        for e in bounds {
                            self.expr(e);
                        }
                    }
                }
            }
            StmtKind::Expr(e) => self.expr(e),
            StmtKind::Return(e) => {
                if let Some(e) = e {
                    self.expr(e);
                }
            }
            StmtKind::Throw { expr, cause } => {
                for e in [expr, cause].into_iter().flatten() {
                    self.expr(e);
                }
            }
            StmtKind::Echo(exprs) => {
                for e in exprs {
                    self.expr(e);
                }
            }
            StmtKind::Assign {
                targets,
                value,
                by_ref,
            } => {
                if *by_ref {
                    if let ExprKind::Ident(name) = &value.kind {
                        self.address(name);
                    }
                }
                for t in targets {
                    self.expr(t);
                }
                self.expr(value);
            }
            StmtKind::CompoundAssign { target, value, .. } => {
                self.expr(target);
                self.expr(value);
            }
            StmtKind::Block(body) => self.block(body),
            StmtKind::NamespaceDecl { body, .. } => self.body(body),
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                self.expr(cond);
                self.block(then_body);
                for (cond, body) in elifs {
                    self.expr(cond);
                    self.block(body);
                }
                if let Some(body) = else_body {
                    self.block(body);
                }
            }
            StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
                self.expr(cond);
                self.block(body);
            }
            StmtKind::For {
                init,
                cond,
                update,
                body,
                ..
            } => {
                self.frames.push(Frame::default());
                if let Some(init) = init {
                    self.stmt(init);
                }
                for e in [cond, update].into_iter().flatten() {
                    self.expr(e);
                }
                self.body(body);
                self.frames.pop();
            }
            StmtKind::ForIn {
                var,
                key,
                iter,
                body,
                else_body,
                ..
            } => {
                self.expr(iter);
                self.frames.push(Frame::default());
                self.bind(var, false);
                if let Some(key) = key {
                    self.bind(key, false);
                }
                self.body(body);
                self.frames.pop();
                if let Some(body) = else_body {
                    self.block(body);
                }
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally,
            } => {
                self.block(body);
                for catch in catches {
                    self.frames.push(Frame::default());
                    for name in [&catch.var_name, &catch.stack_var].into_iter().flatten() {
                        self.bind(name, false);
                    }
                    if let Some(e) = &catch.when_clause {
                        self.expr(e);
                    }
                    self.body(&catch.body);
                    self.frames.pop();
                }
                for body in [else_body, finally].into_iter().flatten() {
                    self.block(body);
                }
            }
            StmtKind::Switch {
                expr,
                cases,
                default,
            } => {
                self.expr(expr);
                for case in cases {
                    for condition in &case.conditions {
                        match condition {
                            CaseCondition::Value(e) | CaseCondition::Comparison { expr: e, .. } => {
                                self.expr(e)
                            }
                            CaseCondition::Range { from, to } => {
                                self.expr(from);
                                self.expr(to);
                            }
                        }
                    }
                    self.block(&case.body);
                }
                if let Some(body) = default {
                    self.block(body);
                }
            }
            StmtKind::Using {
                var,
                resource,
                body,
            } => {
                self.expr(resource);
                self.frames.push(Frame::default());
                self.bind(var, false);
                self.body(body);
                self.frames.pop();
            }
            StmtKind::Lock { expr, body } => {
                self.expr(expr);
                self.block(body);
            }
            StmtKind::Labeled { body, .. } => self.stmt(body),
            StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. }
            | StmtKind::ModuleDecl { members, .. } => {
                for member in members {
                    match member {
                        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                            self.stmt(stmt)
                        }
                        ClassMember::Constructor { params, body, .. } => {
                            self.function(params, body)
                        }
                        ClassMember::Field { init: Some(e), .. } => self.expr(e),
                        ClassMember::Property { getter, setter, .. } => {
                            if let Some(body) = getter {
                                self.function(&[], body);
                            }
                            if let Some(setter) = setter {
                                self.function(std::slice::from_ref(&setter.param), &setter.body);
                            }
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    fn expr(&mut self, expr: &Expression) {
        match &expr.kind {
            ExprKind::Unary {
                op: UnaryOp::AddrOf,
                expr,
            } => {
                if let ExprKind::Ident(name) = &expr.kind {
                    self.address(name);
                }
                self.expr(expr);
            }
            ExprKind::RefOf(place) => match place.as_ref() {
                PlaceExpr::Ident(name) => self.address(name),
                PlaceExpr::Member { object, .. } | PlaceExpr::Deref(object) => self.expr(object),
                PlaceExpr::Index { object, index, .. } => {
                    self.expr(object);
                    self.expr(index);
                }
            },
            ExprKind::Unary { expr, .. }
            | ExprKind::Cast { expr, .. }
            | ExprKind::RefLoad(expr) => self.expr(expr),
            ExprKind::Binary { left, right, .. } => {
                self.expr(left);
                self.expr(right);
            }
            ExprKind::Call { callee, args, .. }
            | ExprKind::New {
                class: callee,
                args,
            } => {
                self.expr(callee);
                for arg in args {
                    if arg.by_ref {
                        if let ExprKind::Ident(name) = &arg.value.kind {
                            self.address(name);
                        }
                    }
                    self.expr(&arg.value);
                }
            }
            ExprKind::Member { object, .. } => self.expr(object),
            ExprKind::Index { object, index, .. } => {
                self.expr(object);
                self.expr(index);
            }
            ExprKind::Ternary { cond, then, else_ } => {
                self.expr(cond);
                self.expr(then);
                self.expr(else_);
            }
            ExprKind::Assign { target, value } => {
                self.expr(target);
                self.expr(value);
            }
            ExprKind::Array(items) => {
                for item in items {
                    if item.by_ref {
                        if let ExprKind::Ident(name) = &item.value.kind {
                            self.address(name);
                        }
                    }
                    if let Some(key) = &item.key {
                        self.expr(key);
                    }
                    self.expr(&item.value);
                }
            }
            ExprKind::Object(props) => {
                for prop in props {
                    match prop {
                        ObjectProperty::KeyValue { value, .. } | ObjectProperty::Spread(value) => {
                            self.expr(value)
                        }
                        ObjectProperty::Computed { key, value } => {
                            self.expr(key);
                            self.expr(value);
                        }
                        ObjectProperty::Method { value, .. }
                        | ObjectProperty::Accessor { value, .. } => self.stmt(value),
                        _ => {}
                    }
                }
            }
            ExprKind::Sequence(exprs) | ExprKind::Tuple(exprs) | ExprKind::Set(exprs) => {
                for e in exprs {
                    self.expr(e);
                }
            }
            ExprKind::FunctionExpr(stmt) => self.stmt(stmt),
            ExprKind::Lambda { params, body, .. } => match body {
                LambdaBody::Block(body) => self.function(params, body),
                LambdaBody::Expr(expr) => {
                    self.frames.push(Frame {
                        function: true,
                        ..Frame::default()
                    });
                    for param in params {
                        self.bind(&param.name, false);
                    }
                    self.expr(expr);
                    self.frames.pop();
                }
            },
            _ => {}
        }
    }
}
