//! Register the `.NET` BCL class wrappers against the user's compiler.
//!
//! Walks `compiler_common::dotnet::class_exports::dotnet_class_exports()`
//! in declared order (which is also inheritance order — `Object` first,
//! leaves last) and, for each wrapper-backed class:
//!
//! 1. Adds `vybe:gui::controlSetProperty` to `chunks[0]` imports (once,
//!    deduped by `add_import`).
//! 2. For concrete leaves (those with `widget_host_fn = Some(_)`), adds
//!    the `vybe:gui::new_<Type>` import to `chunks[0]`.
//! 3. Builds and pushes a setter chunk per property at this class level.
//! 4. Builds and pushes the constructor chunk, referencing the setter
//!    chunk indices via `ref_func` and the parent class via `global_get`
//!    on the parent's name (which the orchestrator already installed in
//!    a previous iteration).
//! 5. Emits `ref_func ctor + global_set <ClassName>` into the script
//!    chunk so user code can `Inherits ClassName`. Also emits a
//!    lowercase alias for VB / Pascal case-insensitivity.
//! 6. Records the class in `defined_classes`, `defined_globals`, and
//!    `pending_classes` so the rest of the compiler treats it as known.
//!
//! Order matters: parent classes must be installed before children
//! because the child's ctor body emits `global_get parent` at install
//! time. The static class table in `compiler_common` is already in the
//! right order (`Object → MarshalByRefObject → Component → Control →
//! ScrollableControl → ContainerControl → Form → …`), so we just iterate
//! it linearly.

use crate::compiler::Compiler;
use crate::emitter as common;
use common::dotnet::class_exports::dotnet_class_exports;
use common::dotnet::classes::builder::{MethodBinding, SetterBinding};
use common::dotnet::classes::{DotnetClass, MethodTarget, builder};

impl Compiler {
    /// Register every `.NET` class wrapper as a callable global on the
    /// current compiler. Called from `compile()` when
    /// `profile.namespaces.use_dotnet` is true.
    pub(crate) fn register_dotnet_classes(&mut self) -> Result<(), String> {
        // Step 1: shared `controlSetProperty` import. All setter chunks call
        // through this single import index — `add_import` dedupes for us, so
        // calling it before each setter is safe. The wrapper chunks stay
        // free of LOCAL imports (strings are pool constants there), so the
        // baked indices always resolve through the script-table fallback.
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

        for export in dotnet_class_exports() {
            let Some(class) = export.wrapper.as_ref() else {
                continue;
            };
            self.register_one_dotnet_class(
                class,
                set_prop_idx,
                get_prop_idx,
                new_controls_collection_idx,
                new_components_collection_idx,
            )?;
        }

        Ok(())
    }

    fn register_one_dotnet_class(
        &mut self,
        class: &DotnetClass,
        set_prop_idx: u16,
        get_prop_idx: u16,
        new_controls_collection_idx: u16,
        new_components_collection_idx: u16,
    ) -> Result<(), String> {
        // ── Step 1: build & push setter chunks for this class's properties ──
        let mut setter_bindings: Vec<SetterBinding<'static>> =
            Vec::with_capacity(class.properties.len());
        let mut getter_bindings: Vec<builder::GetterBinding<'static>> =
            Vec::with_capacity(class.properties.len());
        for prop in class.properties {
            let setter_chunk = builder::build_setter_chunk(class.name, prop, set_prop_idx);
            self.chunks_mut().push(setter_chunk);
            let setter_idx = self.chunks_mut().len() - 1;
            setter_bindings.push(SetterBinding {
                prop_pascal: *prop,
                setter_chunk_idx: setter_idx,
            });

            let getter_chunk = builder::build_getter_chunk(class.name, prop, get_prop_idx);
            self.chunks_mut().push(getter_chunk);
            let getter_idx = self.chunks_mut().len() - 1;
            getter_bindings.push(builder::GetterBinding {
                prop_pascal: *prop,
                getter_chunk_idx: getter_idx,
            });
        }

        // ── Step 2: build & push method thunk chunks for this class ────────
        //
        // Each method thunk forwards `(this, args...)` to either a host
        // import or a dotnet class ctor, depending on `method.target`. For
        // `Host` targets we pre-resolve the import index here so the
        // builder doesn't need to touch the imports vec. For `DotnetCtor`
        // targets the builder uses `global_get` directly and the import
        // index is unused (we pass 0 as a dummy).
        //
        // Method bindings are stored under their lowercased name because
        // that's what the canonical AST emits for `obj.MethodName(...)`.
        let mut method_lowered_names: Vec<String> = Vec::with_capacity(class.methods.len());
        let mut method_thunk_indices: Vec<usize> = Vec::with_capacity(class.methods.len());
        for method in class.methods {
            let (import_idx, body_imports) = match method.target {
                MethodTarget::Host { module, fn_name } => {
                    (self.chunks_mut()[0].add_import(module, fn_name), Vec::new())
                }
                MethodTarget::DotnetCtor { .. } => (0u16, Vec::new()),
                MethodTarget::Body(ops) => {
                    // Pre-resolve every CallHost target's import index
                    // in encounter order. The builder consumes them via
                    // a cursor as it walks the body ops.
                    let targets = builder::collect_body_call_targets(ops);
                    let mut imports = Vec::with_capacity(targets.len());
                    for (module, fn_name) in targets {
                        imports.push(self.chunks_mut()[0].add_import(module, fn_name));
                    }
                    (0u16, imports)
                }
            };
            let thunk =
                builder::build_method_thunk_chunk(class.name, method, import_idx, &body_imports);
            self.chunks_mut().push(thunk);
            method_thunk_indices.push(self.chunks_mut().len() - 1);
            method_lowered_names.push(method.name.to_lowercase());
        }
        let method_bindings: Vec<MethodBinding> = method_lowered_names
            .iter()
            .zip(method_thunk_indices.iter())
            .map(|(name, idx)| MethodBinding {
                method_name: name.as_str(),
                thunk_chunk_idx: *idx,
            })
            .collect();

        // ── Step 3: import the backing host fn if this is a concrete leaf ──
        let widget_new_idx = class
            .widget_host_fn
            .map(|host_fn| self.chunks_mut()[0].add_import(class.widget_host_module, host_fn));

        // ── Step 4: build & push the constructor chunk ─────────────────────
        let ctor_chunk = builder::build_constructor_chunk(
            class,
            &setter_bindings,
            &getter_bindings,
            &method_bindings,
            widget_new_idx,
            new_controls_collection_idx,
            new_components_collection_idx,
        );
        self.chunks_mut().push(ctor_chunk);
        let ctor_idx = self.chunks_mut().len() - 1;

        // ── Step 5: install as a callable global in the script chunk ──────
        let line = self.current_line();
        builder::emit_install_class_global(&mut self.chunks_mut()[0], class.name, ctor_idx, line);

        // ── Step 6: register the class in the compiler's bookkeeping ──────
        // The lowercase form is what the canonical AST uses for identifier
        // lookups in case-insensitive languages (VB, Pascal). The
        // PascalCase form is what case-sensitive languages (C#, Dart) use.
        // Inserting both keeps `defined_globals.contains(name)` honest for
        // either flavour.
        let pascal = class.name.to_string();
        let lower = pascal.to_lowercase();
        self.note_defined_global(&pascal);
        self.note_defined_global(&lower);
        self.note_defined_class(&pascal);
        self.note_defined_class(&lower);

        // For child-class field resolution: record the parent so user
        // subclasses can walk up the chain. Empty `fields` because dotnet
        // classes use setter dispatch, not direct field access.
        let parent = class.parent.map(|p| p.to_lowercase());
        self.note_pending_class(&lower, parent);

        Ok(())
    }
}
