use super::*;

impl Compiler {
    fn is_known_gui_event_name(&self, event: &str) -> bool {
        matches!(
            self.canon(event).as_str(),
            "click"
                | "dblclick"
                | "doubleclick"
                | "load"
                | "unload"
                | "change"
                | "textchanged"
                | "selectedindexchanged"
                | "checkedchanged"
                | "valuechanged"
                | "keypress"
                | "keydown"
                | "keyup"
                | "mousedown"
                | "mouseup"
                | "mousemove"
                | "mouseenter"
                | "mouseleave"
                | "gotfocus"
                | "lostfocus"
                | "enter"
                | "leave"
                | "resize"
                | "paint"
                | "formclosing"
        )
    }

    /// Static cases (push a string constant):
    ///   - `Ident("btn")`          -> "btn"
    ///   - `Me` / `This`           -> current class name (lowercased)
    ///   - `Member { Me, "btn" }` -> "btn"
    /// Dynamic fallback (runtime lookup): compile expr, then read `__control_name`.
    fn emit_event_control_key(&mut self, control: &Expression, line: u32) -> Result<(), String> {
        let is_self_ident = |c: &Compiler, n: &str| {
            let cn = c.canon(n);
            cn == c.profile.self_keyword || cn == "me" || cn == "this" || cn == "mybase"
        };
        let key: Option<String> = match &control.kind {
            ExprKind::This | ExprKind::Super => self.current_class.clone().map(|c| self.canon(&c)),
            ExprKind::Ident(name) if is_self_ident(self, name) => {
                self.current_class.clone().map(|c| self.canon(&c))
            }
            ExprKind::Ident(name) => {
                let is_class_field = if let Some(ref cn) = self.current_class {
                    self.pending_classes
                        .get(cn.as_str())
                        .map(|pc| pc.fields.iter().any(|f| f.eq_ignore_ascii_case(name)))
                        .unwrap_or(false)
                } else {
                    false
                };
                if is_class_field {
                    Some(self.canon(name))
                } else {
                    None
                }
            }
            ExprKind::Member { object, field, .. } => {
                let is_self = matches!(&object.kind, ExprKind::This | ExprKind::Super)
                    || matches!(&object.kind, ExprKind::Ident(n) if is_self_ident(self, n));
                if is_self {
                    Some(self.canon(field))
                } else {
                    None
                }
            }
            _ => None };
        if let Some(k) = key {
            self.emit_const(Value::String(Arc::from(k.as_str())));
        } else {
            self.compile_expr(control)?;
            common::gui::emit_get_control_name(self.chunk(), line);
        }
        Ok(())
    }

    fn build_event_binding_target(&self, control: &Expression, event: &str) -> Expression {
        let control = match &control.kind {
            ExprKind::Cast { expr, .. } => expr.as_ref(),
            _ => control };
        if event.is_empty() {
            control.clone()
        } else {
            Expression::new(ExprKind::Member {
                object: Box::new(control.clone()),
                field: event.to_string(),
                null_safe: false })
        }
    }

    fn event_receiver_type_hint(&self, control: &Expression) -> Option<String> {
        let is_self_ident = |c: &Compiler, n: &str| {
            let cn = c.canon(n);
            cn == c.profile.self_keyword || cn == "me" || cn == "this" || cn == "mybase"
        };

        match &control.kind {
            ExprKind::This | ExprKind::Super => self.current_class.clone(),
            ExprKind::Ident(name) if is_self_ident(self, name) => self.current_class.clone(),
            _ => self.infer_expr_type_hint(control) }
    }

    fn type_uses_winforms_event_host(&self, candidate_type: &str) -> bool {
        // A concrete GUI control (`Button`, `Form`, `TextBox`, …) is a WinForms
        // Control even though it no longer registers a pending class (control
        // ctor globals are retired). Recognize it via the shared control-name
        // table so `btn.Click += handler` binds to the widget (`onEvent`)
        // rather than only combining an in-memory delegate.
        if !common::gui::canonical_control_name(candidate_type).is_empty() {
            return true;
        }
        if let Some(mut current) = self.resolve_pending_class_name_for_type_hint(candidate_type) {
            let mut visited = std::collections::HashSet::new();
            loop {
                let current_key = self.canon(&current);
                if !visited.insert(current_key.clone()) {
                    return false;
                }

                if current.eq_ignore_ascii_case("Control")
                    || current.eq_ignore_ascii_case("Form")
                    || current.eq_ignore_ascii_case("System.Windows.Forms.Control")
                    || current.eq_ignore_ascii_case("System.Windows.Forms.Form")
                {
                    return true;
                }

                let Some(parent) = self
                    .pending_classes
                    .get(&current_key)
                    .and_then(|pending| pending.parent.clone())
                else {
                    return false;
                };
                if !self.pending_classes.contains_key(&self.canon(&parent)) {
                    return self
                        .reflection_is_assignable_from("System.Windows.Forms.Control", &parent);
                }
                current = parent;
            }
        }

        self.reflection_is_assignable_from("System.Windows.Forms.Control", candidate_type)
    }

    fn should_use_gui_event_host(&self, control: &Expression, event: &str) -> bool {
        if event.is_empty() {
            return false;
        }

        // An `object`-declared receiver is a receiver whose type is UNKNOWN,
        // not a receiver of a particular language. The deciding signal is the
        // EVENT NAME being one the shared GUI surface declares — a control
        // wired to `Click`/`TextChanged` is a control whoever spelled it. The
        // language gate that used to lead here answered nothing this test
        // does not.
        if self.is_known_gui_event_name(event)
            && self
                .event_receiver_type_hint(control)
                .as_deref()
                .is_some_and(|type_hint| {
                    matches!(
                        Self::normalize_type_hint(type_hint).as_str(),
                        "object" | "system.object"
                    )
                })
        {
            return true;
        }

        self.event_receiver_type_hint(control)
            .map(|type_hint| self.type_uses_winforms_event_host(&type_hint))
            .unwrap_or(false)
    }

    fn compile_delegate_event_invoke(
        &mut self,
        target: &Expression,
        args: &[Expression],
    ) -> Result<(), String> {
        // A raised event may hold a multicast delegate (array of handlers).
        // The shared invoker iterates and calls each, and yields null for a
        // null delegate — so no explicit null guard is needed here.
        self.compile_expr(target)?;
        for arg in args {
            self.compile_expr(arg)?;
        }
        common::delegates::emit_invoke(
            &mut self.chunks,
            self.current,
            (args.len() + 1) as u8,
            self.line,
        );
        self.emit(Op::DROP);
        Ok(())
    }

    pub(super) fn compile_add_handler_stmt(
        &mut self,
        control: &Expression,
        event: &str,
        handler: &Expression,
    ) -> Result<(), String> {
        if self.should_use_gui_event_host(control, event) {
            let line = self.line;
            let bind_idx = self.import("vybe:gui", common::gui::HOST_FN_BIND_EVENT);
            self.emit_event_control_key(control, line)?;
            self.emit_const(Value::String(Arc::from(event)));
            self.compile_expr(handler)?;
            common::gui::emit_bind_event(self.chunk(), bind_idx, line);
            self.emit(Op::DROP);
            return Ok(());
        }

        let event_target = self.build_event_binding_target(control, event);
        self.compile_expr(&event_target)?;
        self.compile_expr(handler)?;
        common::delegates::emit_combine(&mut self.chunks, self.current, self.line);
        self.compile_assign_target(&event_target)
    }

    pub(super) fn compile_remove_handler_stmt(
        &mut self,
        control: &Expression,
        event: &str,
        handler: &Expression,
    ) -> Result<(), String> {
        if self.should_use_gui_event_host(control, event) {
            let line = self.line;
            let unbind_idx = self.import("vybe:gui", common::gui::HOST_FN_UNBIND_EVENT);
            self.emit_event_control_key(control, line)?;
            self.emit_const(Value::String(Arc::from(event)));
            self.compile_expr(handler)?;
            common::gui::emit_unbind_event(self.chunk(), unbind_idx, line);
            self.emit(Op::DROP);
            return Ok(());
        }

        let event_target = self.build_event_binding_target(control, event);
        self.compile_expr(&event_target)?;
        self.compile_expr(handler)?;
        common::delegates::emit_remove(&mut self.chunks, self.current, self.line);
        self.compile_assign_target(&event_target)
    }

    pub(super) fn compile_raise_event_stmt(
        &mut self,
        event_name: &str,
        args: &[Expression],
    ) -> Result<(), String> {
        let event_field = self.canon(event_name);
        let target = if self.current_member_is_static {
            if let Some(class_name) = self.current_class.as_deref() {
                Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(class_name)),
                    field: event_field,
                    null_safe: false })
            } else {
                Expression::ident(&event_field)
            }
        } else if self.current_class.is_some() {
            Expression::new(ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::This)),
                field: event_field,
                null_safe: false })
        } else {
            Expression::ident(&event_field)
        };
        self.compile_delegate_event_invoke(&target, args)
    }
}

// ── Chunk-level events emit ────────────────────────────────────────────
// Free functions over `&mut Chunk`, merged in from the former `emitter::events`
// module. The `impl Compiler` walkers above and these primitives are the two
// halves of the same topic and now live in one file.
use vybe_ast::{BinOp, ExprKind, Expression, StmtKind};
// Event-handler AST lowering — language-agnostic helpers that normalise
// event-subscription syntax (`AddHandler`, `control.Click += handler`, …) onto
// the canonical `StmtKind::AddHandler`/`RemoveHandler` nodes. Kept separate
// from `gui.rs` because events are not necessarily a GUI concern.
//
// The emit side lives in the compiler (`primitives/events.rs`) and the shared
// `vybe:gui` binding in `gui.rs`; these are the *walker-facing* builders any
// front-end can reuse.

pub fn add_handler_stmt(
    control: Expression,
    event: impl Into<String>,
    handler: Expression,
) -> StmtKind {
    StmtKind::AddHandler {
        control,
        event: event.into(),
        handler }
}

pub fn remove_handler_stmt(
    control: Expression,
    event: impl Into<String>,
    handler: Expression,
) -> StmtKind {
    StmtKind::RemoveHandler {
        control,
        event: event.into(),
        handler }
}

pub fn lower_event_compound_assignment(expr: &Expression) -> Option<StmtKind> {
    let ExprKind::Assign { target, value } = &expr.kind else {
        return None;
    };
    let ExprKind::Member {
        object: ev_obj,
        field: ev_field,
        ..
    } = &target.kind
    else {
        return None;
    };
    let ExprKind::Binary { op, left, right } = &value.kind else {
        return None;
    };

    let same_target = matches!(
        &left.kind,
        ExprKind::Member { object, field, .. } if member_eq(object, field, ev_obj, ev_field)
    );
    let handler = unwrap_event_handler(right)?;

    if !same_target {
        return None;
    }

    let event_name = ev_field.to_lowercase();
    let control = (**ev_obj).clone();
    Some(match op {
        BinOp::Add => add_handler_stmt(control, event_name, handler.clone()),
        BinOp::Sub => remove_handler_stmt(control, event_name, handler.clone()),
        _ => return None })
}

fn unwrap_event_handler(expr: &Expression) -> Option<&Expression> {
    if is_event_handler_expr(expr) {
        return Some(expr);
    }

    match &expr.kind {
        ExprKind::New { args, .. } if args.len() == 1 => {
            let inner = &args[0].value;
            if is_event_handler_expr(inner) {
                Some(inner)
            } else {
                None
            }
        }
        _ => None }
}

pub fn is_event_handler_expr(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Ident(_) | ExprKind::Member { .. } | ExprKind::Lambda { .. }
    )
}

/// True if `field` is a known WinForms/GUI event name (case-insensitive).
/// Used by language walkers to recognise `control.Click += handler` as event
/// subscription rather than ordinary numeric compound-assignment.
pub fn is_known_gui_event_field(field: &str) -> bool {
    matches!(
        field.to_lowercase().as_str(),
        "click"
            | "dblclick"
            | "doubleclick"
            | "load"
            | "unload"
            | "change"
            | "textchanged"
            | "selectedindexchanged"
            | "checkedchanged"
            | "valuechanged"
            | "keypress"
            | "keydown"
            | "keyup"
            | "mousedown"
            | "mouseup"
            | "mousemove"
            | "mouseenter"
            | "mouseleave"
            | "gotfocus"
            | "lostfocus"
            | "enter"
            | "leave"
            | "resize"
            | "paint"
            | "formclosing"
    )
}

fn member_eq(obj_a: &Expression, field_a: &str, obj_b: &Expression, field_b: &str) -> bool {
    if !field_a.eq_ignore_ascii_case(field_b) {
        return false;
    }

    match (&obj_a.kind, &obj_b.kind) {
        (ExprKind::Ident(a), ExprKind::Ident(b)) => a == b,
        (ExprKind::This, ExprKind::This) => true,
        (
            ExprKind::Member {
                object: inner_a,
                field: inner_field_a,
                ..
            },
            ExprKind::Member {
                object: inner_b,
                field: inner_field_b,
                ..
            },
        ) => member_eq(inner_a, inner_field_a, inner_b, inner_field_b),
        _ => false }
}
