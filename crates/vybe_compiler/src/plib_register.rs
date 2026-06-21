//! Register Pascal library adapters against the compiler.

use crate::ast::{Argument, ClassMember, ExprKind, Expression, ImportKind, Module, StmtKind};
use crate::compiler::Compiler;
use crate::emitter as common;
use crate::platforms::plib::gcl::{GclClass, GclMethodTarget, builder, gcl_classes, is_gcl_unit};

pub(crate) fn module_uses_plib_gcl(module: &Module) -> bool {
    module.imports.iter().any(|import| match &import.kind {
        ImportKind::Simple { path, .. }
        | ImportKind::Named { path, .. }
        | ImportKind::Wildcard { path, .. }
        | ImportKind::Default { path, .. } => is_gcl_unit(path),
    }) || module.body.iter().any(stmt_uses_gcl_dialog)
}

impl Compiler {
    pub(crate) fn register_plib_gcl_classes(&mut self) -> Result<(), String> {
        let set_prop_idx = self.chunks_mut()[0]
            .add_import(common::gui::GUI_MODULE, common::gui::HOST_FN_SET_PROPERTY);
        let get_prop_idx = self.chunks_mut()[0]
            .add_import(common::gui::GUI_MODULE, common::gui::HOST_FN_GET_PROPERTY);
        let new_controls_collection_idx = self.chunks_mut()[0].add_import(
            common::gui::GUI_MODULE,
            common::gui::HOST_FN_NEW_CONTROLS_COLLECTION,
        );
        let new_components_collection_idx = self.chunks_mut()[0].add_import(
            common::gui::GUI_MODULE,
            common::gui::HOST_FN_NEW_COMPONENTS_COLLECTION,
        );
        let bind_event_idx = self.chunks_mut()[0]
            .add_import(common::gui::GUI_MODULE, common::gui::HOST_FN_BIND_EVENT);
        let run_application_idx = self.chunks_mut()[0]
            .add_import(common::gui::GUI_MODULE, common::gui::HOST_FN_RUN_APPLICATION);
        let app_exit_idx = self.chunks_mut()[0]
            .add_import(common::gui::GUI_MODULE, common::gui::HOST_FN_APP_EXIT);
        let msg_box_idx = self.chunks_mut()[0].add_import(common::gui::GUI_MODULE, "msgBox");

        for class in gcl_classes() {
            self.register_one_plib_gcl_class(
                class,
                set_prop_idx,
                get_prop_idx,
                bind_event_idx,
                run_application_idx,
                new_controls_collection_idx,
                new_components_collection_idx,
            )?;
        }

        let run_chunk = builder::build_application_run_chunk(run_application_idx);
        self.chunks_mut().push(run_chunk);
        let run_chunk_idx = self.chunks_mut().len() - 1;
        let exit_chunk = builder::build_application_exit_chunk(app_exit_idx);
        self.chunks_mut().push(exit_chunk);
        let exit_chunk_idx = self.chunks_mut().len() - 1;
        let initialize_chunk = builder::build_application_initialize_chunk();
        self.chunks_mut().push(initialize_chunk);
        let initialize_chunk_idx = self.chunks_mut().len() - 1;
        let title_setter_chunk = builder::build_application_title_setter_chunk();
        self.chunks_mut().push(title_setter_chunk);
        let title_setter_idx = self.chunks_mut().len() - 1;
        let title_getter_chunk = builder::build_application_title_getter_chunk();
        self.chunks_mut().push(title_getter_chunk);
        let title_getter_idx = self.chunks_mut().len() - 1;
        let line = self.current_line();
        builder::emit_install_application_global(
            &mut self.chunks_mut()[0],
            run_chunk_idx,
            exit_chunk_idx,
            initialize_chunk_idx,
            title_setter_idx,
            title_getter_idx,
            line,
        );
        self.note_defined_global("Application");
        self.note_defined_global("application");

        let show_message = builder::build_show_message_chunk(msg_box_idx);
        self.chunks_mut().push(show_message);
        let show_message_idx = self.chunks_mut().len() - 1;
        let message_dlg = builder::build_message_dlg_chunk(msg_box_idx);
        self.chunks_mut().push(message_dlg);
        let message_dlg_idx = self.chunks_mut().len() - 1;
        let line = self.current_line();
        builder::emit_install_function_global(
            &mut self.chunks_mut()[0],
            "ShowMessage",
            show_message_idx,
            line,
        );
        builder::emit_install_function_global(
            &mut self.chunks_mut()[0],
            "MessageDlg",
            message_dlg_idx,
            line,
        );
        self.note_defined_global("ShowMessage");
        self.note_defined_global("showmessage");
        self.note_defined_global("MessageDlg");
        self.note_defined_global("messagedlg");
        self.defined_functions.insert("ShowMessage".to_string());
        self.defined_functions.insert("showmessage".to_string());
        self.defined_functions.insert("MessageDlg".to_string());
        self.defined_functions.insert("messagedlg".to_string());

        Ok(())
    }

    fn register_one_plib_gcl_class(
        &mut self,
        class: &GclClass,
        set_prop_idx: u16,
        get_prop_idx: u16,
        bind_event_idx: u16,
        run_application_idx: u16,
        new_controls_collection_idx: u16,
        new_components_collection_idx: u16,
    ) -> Result<(), String> {
        let mut setter_bindings = Vec::with_capacity(class.properties.len());
        let mut getter_bindings = Vec::with_capacity(class.properties.len());
        for property in class.properties {
            let setter_idx = if builder::is_event_property(property) {
                bind_event_idx
            } else {
                set_prop_idx
            };
            let size_sync_idx = if matches!(
                property.to_ascii_lowercase().as_str(),
                "clientwidth" | "clientheight"
            ) {
                Some(run_application_idx)
            } else {
                None
            };
            let setter = builder::build_setter_chunk(class.name, property, setter_idx, size_sync_idx);
            self.chunks_mut().push(setter);
            setter_bindings.push(builder::AccessorBinding {
                property_name: property,
                chunk_idx: self.chunks_mut().len() - 1,
            });

            let getter = builder::build_getter_chunk(class.name, property, get_prop_idx);
            self.chunks_mut().push(getter);
            getter_bindings.push(builder::AccessorBinding {
                property_name: property,
                chunk_idx: self.chunks_mut().len() - 1,
            });
        }

        let mut method_names = Vec::with_capacity(class.methods.len());
        let mut method_indices = Vec::with_capacity(class.methods.len());
        for method in class.methods {
            let import_idx = match method.target {
                GclMethodTarget::Host { module, fn_name } => {
                    self.chunks_mut()[0].add_import(module, fn_name)
                }
            };
            let thunk = builder::build_method_chunk(class.name, method, import_idx);
            self.chunks_mut().push(thunk);
            method_names.push(method.name.to_string());
            method_indices.push(self.chunks_mut().len() - 1);
            let lower = method.name.to_lowercase();
            if lower != method.name {
                method_names.push(lower);
                method_indices.push(self.chunks_mut().len() - 1);
            }
        }
        let method_bindings: Vec<builder::MethodBinding> = method_names
            .iter()
            .zip(method_indices.iter())
            .map(|(name, idx)| builder::MethodBinding {
                method_name: name.as_str(),
                chunk_idx: *idx,
            })
            .collect();

        let widget_new_idx = class
            .widget_host_fn
            .map(|host_fn| self.chunks_mut()[0].add_import(common::gui::GUI_MODULE, host_fn));

        let ctor = builder::build_constructor_chunk(
            class,
            &setter_bindings,
            &getter_bindings,
            &method_bindings,
            widget_new_idx,
            new_controls_collection_idx,
            new_components_collection_idx,
        );
        self.chunks_mut().push(ctor);
        let ctor_idx = self.chunks_mut().len() - 1;

        let line = self.current_line();
        builder::emit_install_class_global(&mut self.chunks_mut()[0], class.name, ctor_idx, line);

        let pascal = class.name.to_string();
        let lower = pascal.to_lowercase();
        self.note_defined_global(&pascal);
        self.note_defined_global(&lower);
        self.note_defined_class(&pascal);
        self.note_defined_class(&lower);
        self.note_pending_class(&lower, class.parent.map(|parent| parent.to_lowercase()));

        Ok(())
    }
}

fn stmt_uses_gcl_dialog(stmt: &crate::ast::Statement) -> bool {
    match &stmt.kind {
        StmtKind::Expr(expr) => expr_uses_gcl_dialog(expr),
        StmtKind::Block(stmts) => stmts.iter().any(stmt_uses_gcl_dialog),
        StmtKind::VarDecl { declarations, .. } => declarations
            .iter()
            .any(|decl| decl.init.as_ref().is_some_and(expr_uses_gcl_dialog)),
        StmtKind::FunctionDecl { body, .. } => body.iter().any(stmt_uses_gcl_dialog),
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => members.iter().any(member_uses_gcl_dialog),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            expr_uses_gcl_dialog(cond)
                || then_body.iter().any(stmt_uses_gcl_dialog)
                || elifs
                    .iter()
                    .any(|(cond, body)| expr_uses_gcl_dialog(cond) || body.iter().any(stmt_uses_gcl_dialog))
                || else_body
                    .as_ref()
                    .is_some_and(|body| body.iter().any(stmt_uses_gcl_dialog))
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            init.as_ref().is_some_and(|stmt| stmt_uses_gcl_dialog(stmt))
                || cond.as_ref().is_some_and(expr_uses_gcl_dialog)
                || update.as_ref().is_some_and(expr_uses_gcl_dialog)
                || body.iter().any(stmt_uses_gcl_dialog)
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            expr_uses_gcl_dialog(iter)
                || body.iter().any(stmt_uses_gcl_dialog)
                || else_body
                    .as_ref()
                    .is_some_and(|body| body.iter().any(stmt_uses_gcl_dialog))
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            expr_uses_gcl_dialog(cond)
                || body.iter().any(stmt_uses_gcl_dialog)
                || else_body
                    .as_ref()
                    .is_some_and(|body| body.iter().any(stmt_uses_gcl_dialog))
        }
        StmtKind::DoWhile { body, cond, .. } => {
            expr_uses_gcl_dialog(cond) || body.iter().any(stmt_uses_gcl_dialog)
        }
        StmtKind::Assign { targets, value } => {
            targets.iter().any(expr_uses_gcl_dialog) || expr_uses_gcl_dialog(value)
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            expr_uses_gcl_dialog(target) || expr_uses_gcl_dialog(value)
        }
        _ => false,
    }
}

fn member_uses_gcl_dialog(member: &ClassMember) -> bool {
    match member {
        ClassMember::Method(stmt) => stmt_uses_gcl_dialog(stmt),
        ClassMember::Constructor { body, .. } => body.iter().any(stmt_uses_gcl_dialog),
        ClassMember::Field { init, .. } => init.as_ref().is_some_and(expr_uses_gcl_dialog),
        ClassMember::Property { getter, setter, .. } => {
            getter
                .as_ref()
                .is_some_and(|body| body.iter().any(stmt_uses_gcl_dialog))
                || setter
                    .as_ref()
                    .is_some_and(|setter| setter.body.iter().any(stmt_uses_gcl_dialog))
        }
        _ => false,
    }
}

fn expr_uses_gcl_dialog(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Call { callee, args, .. } => {
            matches!(
                &callee.kind,
                ExprKind::Ident(name)
                    if name.eq_ignore_ascii_case("ShowMessage")
                        || name.eq_ignore_ascii_case("MessageDlg")
            ) || expr_uses_gcl_dialog(callee)
                || args.iter().any(arg_uses_gcl_dialog)
        }
        ExprKind::Member { object, .. } => expr_uses_gcl_dialog(object),
        ExprKind::Binary { left, right, .. } => {
            expr_uses_gcl_dialog(left) || expr_uses_gcl_dialog(right)
        }
        ExprKind::Unary { expr, .. } => expr_uses_gcl_dialog(expr),
        ExprKind::Assign { target, value } => {
            expr_uses_gcl_dialog(target) || expr_uses_gcl_dialog(value)
        }
        ExprKind::Array(items) => items.iter().any(|item| {
            item.key.as_ref().is_some_and(expr_uses_gcl_dialog)
                || expr_uses_gcl_dialog(&item.value)
        }),
        _ => false,
    }
}

fn arg_uses_gcl_dialog(arg: &Argument) -> bool {
    expr_uses_gcl_dialog(&arg.value)
}
