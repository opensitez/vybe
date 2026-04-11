//! Register the `.NET` BCL class wrappers against the user's compiler.
//!
//! Walks `compiler_common::dotnet::classes::dotnet_classes()` in declared
//! order (which is also inheritance order — `Object` first, leaves last)
//! and, for each class:
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
use vybe_compiler_common as common;
use common::dotnet::classes::{dotnet_classes, builder, DotnetClass};
use common::dotnet::classes::builder::SetterBinding;

impl Compiler {
    /// Register every `.NET` class wrapper as a callable global on the
    /// current compiler. Called from `compile()` when
    /// `profile.namespaces.use_dotnet` is true.
    pub(crate) fn register_dotnet_classes(&mut self) -> Result<(), String> {
        // Step 1: shared `controlSetProperty` import. All setter chunks call
        // through this single import index — `add_import` dedupes for us, so
        // calling it before each setter is safe.
        let set_prop_idx = self.chunks_mut()[0]
            .add_import(common::gui::GUI_MODULE, common::gui::HOST_FN_SET_PROPERTY);

        for class in dotnet_classes() {
            self.register_one_dotnet_class(class, set_prop_idx)?;
        }

        Ok(())
    }

    fn register_one_dotnet_class(
        &mut self,
        class: &DotnetClass,
        set_prop_idx: u16,
    ) -> Result<(), String> {
        // ── Step 1: build & push setter chunks for this class's properties ──
        let mut bindings: Vec<SetterBinding<'static>> = Vec::with_capacity(class.properties.len());
        for prop in class.properties {
            let setter_chunk = builder::build_setter_chunk(class.name, prop, set_prop_idx);
            self.chunks_mut().push(setter_chunk);
            let setter_idx = self.chunks_mut().len() - 1;
            bindings.push(SetterBinding {
                prop_pascal: *prop,
                setter_chunk_idx: setter_idx,
            });
        }

        // ── Step 2: import vybe:gui::new_<Type> if this is a concrete leaf ──
        let widget_new_idx = class.widget_host_fn.map(|host_fn| {
            self.chunks_mut()[0].add_import(common::gui::GUI_MODULE, host_fn)
        });

        // ── Step 3: build & push the constructor chunk ─────────────────────
        let ctor_chunk = builder::build_constructor_chunk(class, &bindings, widget_new_idx);
        self.chunks_mut().push(ctor_chunk);
        let ctor_idx = self.chunks_mut().len() - 1;

        // ── Step 4: install as a callable global in the script chunk ──────
        let line = self.current_line();
        builder::emit_install_class_global(&mut self.chunks_mut()[0], class.name, ctor_idx, line);

        // ── Step 5: register the class in the compiler's bookkeeping ──────
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
