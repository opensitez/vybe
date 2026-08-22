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

    // `emit_event_control_key` is GONE. It built the NAME key for the
    // name-addressed event registry (`Ident("btn")` -> "btn", `Me` -> the class
    // name, else a runtime `__control_name` read). Both callers now subscribe
    // through the DOM by the `on<event>` role, so there is no name to build —
    // and the key it produced was the empty string for any control created by
    // `createElement`, which is why those subscriptions never fired.

    fn build_event_binding_target(&self, control: &Expression, event: &str) -> Expression {
        let control = match &control.kind {
            ExprKind::Cast { expr, .. } => expr.as_ref(),
            _ => control,
        };
        if event.is_empty() {
            control.clone()
        } else {
            Expression::new(ExprKind::Member {
                object: Box::new(control.clone()),
                field: event.to_string(),
                null_safe: false,
            })
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
            _ => self.infer_expr_type_hint(control),
        }
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
        self.emit_multicast_delegate_invoke(args.len() as u8);
        self.emit(Op::DROP);
        Ok(())
    }

    pub(super) fn compile_add_handler_stmt(
        &mut self,
        control: &Expression,
        event: &str,
        handler: &Expression,
    ) -> Result<(), String> {
        // Subscribing on a DOM control IS `addEventListener`. `OnClick :=
        // handler` — the same subscription spelled as a property — already
        // emits exactly that, so this hands the statement spelling to the same
        // emit rather than keeping a second registry.
        //
        // The one it replaces was name-keyed: `onEvent(control.__control_name,
        // …)` into `GuiState`. An element created by `createElement` has no
        // `__control_name` (the factory that used to assign one is
        // gone), so every `AddHandler`/`Handles`/`+=` subscription registered
        // under an EMPTY key and no button in any WinForms program could fire.
        // Nothing errored — the control existed, the handler existed, and the
        // two were never connected.
        if self.control_receiver_is_element(control) {
            let line = self.line;
            self.compile_expr(control)?;
            self.compile_expr(handler)?;
            // `on` + the event word is the role every frontend already spells:
            // Pascal writes the property `OnClick`, and this is the same role
            // reached from `Click`. `event_role_type` strips it back off, so
            // `Load`/`Timer` register alongside `click` with no table.
            let role = format!("on{}", event.to_ascii_lowercase());
            self.emit_gui_property_set(&role, line);
            self.emit(Op::DROP);
            return Ok(());
        }
        // The receiver is a CONTROL that the static type did not prove is an
        // element — `object`-declared, or a type the frontend flags. It is
        // still a control, so it still subscribes through the DOM: same role,
        // same emit, just reached without the static proof.
        //
        // This used to bind through the name-keyed host instead, which is the
        // broken registry the comment above describes — an element built by
        // `createElement` has no `__control_name`, so every subscription taking
        // this branch registered under an EMPTY key and could never fire.
        // Routing it to the role means the two arms cannot disagree.
        if self.should_use_gui_event_host(control, event) {
            let line = self.line;
            self.compile_expr(control)?;
            self.compile_expr(handler)?;
            let role = format!("on{}", event.to_ascii_lowercase());
            self.emit_gui_property_set(&role, line);
            self.emit(Op::DROP);
            return Ok(());
        }

        let event_target = self.build_event_binding_target(control, event);
        self.compile_expr(&event_target)?;
        self.compile_expr(handler)?;
        common::delegates::emit_combine(&mut self.chunks, self.current, self.line);
        self.compile_assign_target(&event_target)
    }

    /// Is this subscription's receiver a DOM element?
    ///
    /// The receiver's STATIC type answers it, the same question
    /// `emit_control_property_set` asks before sending a property write to the
    /// document — a control is an element for its events exactly when it is one
    /// for its properties.
    fn control_receiver_is_element(&self, control: &Expression) -> bool {
        self.event_receiver_type_hint(control)
            .map(|type_hint| Self::normalize_type_hint(&type_hint))
            .is_some_and(|class_name| self.control_element_for_type(&class_name).is_some())
    }

    pub(super) fn compile_remove_handler_stmt(
        &mut self,
        control: &Expression,
        event: &str,
        handler: &Expression,
    ) -> Result<(), String> {
        // The mirror of `compile_add_handler_stmt`: an element unsubscribes
        // through the DOM, like it subscribes. Without this branch the two
        // halves used different registries and removal was a no-op.
        if self.control_receiver_is_element(control) {
            let line = self.line;
            self.compile_expr(control)?;
            self.compile_expr(handler)?;
            self.emit_remove_event_listener(event, line);
            self.emit(Op::DROP);
            return Ok(());
        }
        // Mirror of the subscribe arm above: an unproven-but-real control
        // unsubscribes through the DOM, exactly as it subscribed. The two
        // halves MUST use the same registry — that is the bug this file's
        // first comment records, and binding through the host here while the
        // element arm removed through the DOM would recreate it one branch
        // over.
        if self.should_use_gui_event_host(control, event) {
            let line = self.line;
            self.compile_expr(control)?;
            self.compile_expr(handler)?;
            self.emit_remove_event_listener(event, line);
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
                    null_safe: false,
                })
            } else {
                Expression::ident(&event_field)
            }
        } else if self.current_class.is_some() {
            Expression::new(ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::This)),
                field: event_field,
                null_safe: false,
            })
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
// binding in `gui.rs`; these are the *walker-facing* builders any front-end
// can reuse.

pub fn add_handler_stmt(
    control: Expression,
    event: impl Into<String>,
    handler: Expression,
) -> StmtKind {
    StmtKind::AddHandler {
        control,
        event: event.into(),
        handler,
    }
}

pub fn remove_handler_stmt(
    control: Expression,
    event: impl Into<String>,
    handler: Expression,
) -> StmtKind {
    StmtKind::RemoveHandler {
        control,
        event: event.into(),
        handler,
    }
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
        _ => return None,
    })
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
        _ => None,
    }
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
        _ => false,
    }
}
