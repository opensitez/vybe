//! Register Pascal library adapters against the compiler.

use crate::ast::{ImportKind, Module};
use crate::compiler::Compiler;
use crate::emitter as common;
use crate::platforms::plib::gcl::{GclClass, GclMethodTarget, builder, gcl_classes, is_gcl_unit};

pub(crate) fn module_uses_plib_gcl(module: &Module) -> bool {
    module.imports.iter().any(|import| match &import.kind {
        ImportKind::Simple { path, .. }
        | ImportKind::Named { path, .. }
        | ImportKind::Wildcard { path, .. }
        | ImportKind::Default { path, .. } => is_gcl_unit(path),
    })
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

        for class in gcl_classes() {
            self.register_one_plib_gcl_class(
                class,
                set_prop_idx,
                get_prop_idx,
                bind_event_idx,
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
        let line = self.current_line();
        builder::emit_install_application_global(
            &mut self.chunks_mut()[0],
            run_chunk_idx,
            exit_chunk_idx,
            line,
        );
        self.note_defined_global("Application");
        self.note_defined_global("application");

        Ok(())
    }

    fn register_one_plib_gcl_class(
        &mut self,
        class: &GclClass,
        set_prop_idx: u16,
        get_prop_idx: u16,
        bind_event_idx: u16,
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
            let setter = builder::build_setter_chunk(class.name, property, setter_idx);
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
            method_names.push(method.name.to_lowercase());
            method_indices.push(self.chunks_mut().len() - 1);
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
