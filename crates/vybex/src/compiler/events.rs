use super::*;

impl Compiler {
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
            _ => None,
        };
        if let Some(k) = key {
            self.emit_const(Value::String(Arc::from(k.as_str())));
        } else {
            self.compile_expr(control)?;
            common::gui::emit_get_control_name(self.chunk(), line);
        }
        Ok(())
    }

    fn build_event_binding_target(&self, control: &Expression, event: &str) -> Expression {
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
                    return self.reflection_is_assignable_from("System.Windows.Forms.Control", &parent);
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

        self.event_receiver_type_hint(control)
            .map(|type_hint| self.type_uses_winforms_event_host(&type_hint))
            .unwrap_or(true)
    }

    fn compile_delegate_event_invoke(&mut self, target: &Expression, args: &[Expression]) -> Result<(), String> {
        let delegate_slot = self.define_local("__raise_event_delegate");
        self.compile_expr(target)?;
        self.emit_u16(Op::LOCAL_SET, delegate_slot);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, delegate_slot);
        self.emit(Op::REF_IS_NULL);
        let has_delegate = self.emit_jump(Op::BR_IF_FALSE);
        let done = self.emit_jump(Op::BR);

        self.patch_jump(has_delegate);
        self.emit_u16(Op::LOCAL_GET, delegate_slot);
        for arg in args {
            self.compile_expr(arg)?;
        }
        self.emit_u8(Op::CALL_REF, args.len() as u8);
        self.emit(Op::DROP);
        self.patch_jump(done);
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
        let target = if self.current_class.is_some() {
            Expression::new(ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::This)),
                field: event_name.to_string(),
                null_safe: false,
            })
        } else {
            Expression::ident(event_name)
        };
        self.compile_delegate_event_invoke(&target, args)
    }
}