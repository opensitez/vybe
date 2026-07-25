//! The vybe platform as a `vybe_bytecode::Plugin` — one plugin, same type as all
//! the others. Its `init` registers everything this provider offers: the
//! always-on `vybe:gui` 2D drawing surface, and (under the `gui` feature, when
//! the `Gui` capability is granted and a `GuiState` is present) the
//! widget-backed form/control/canvas surface. A plugin is whatever capabilities
//! it registers, so a single `init` covers drawing + gui + canvas together.

/// The vybe platform plugin. Owns its widget state when it has one (the
/// gui/canvas surface needs it); a plugin owns the state it registers over —
/// the host never constructs it. Use [`Plugin::new`] (drawing only) or
/// [`Plugin::with_gui`] (creates a fresh `GuiState`), then read the handle back
/// with [`Plugin::gui_state`]. In a dylib this factory becomes state the
/// plugin's own `init` creates, and the accessor a registered handle.
pub struct Plugin {
    #[cfg(feature = "gui")]
    gui: Option<std::sync::Arc<std::sync::Mutex<crate::gui_state::GuiState>>>,
}

impl Default for Plugin {
    fn default() -> Self {
        Self::new()
    }
}

impl Plugin {
    /// A drawing-only vybe plugin (no widget state).
    pub fn new() -> Self {
        Self {
            #[cfg(feature = "gui")]
            gui: None,
        }
    }

    /// A gui-capable vybe plugin that **owns** a freshly created `GuiState`.
    /// The host retrieves the shared handle afterwards via [`gui_state`].
    #[cfg(feature = "gui")]
    pub fn with_gui() -> Self {
        Self {
            gui: Some(std::sync::Arc::new(std::sync::Mutex::new(
                crate::gui_state::GuiState::new(),
            ))),
        }
    }

    /// The shared `GuiState` this plugin owns, if any (for the form launcher).
    #[cfg(feature = "gui")]
    pub fn gui_state(&self) -> Option<std::sync::Arc<std::sync::Mutex<crate::gui_state::GuiState>>> {
        self.gui.clone()
    }
}

impl vybe_bytecode::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "vybe"
    }

    fn init(&self, fw: &mut vybe_bytecode::Framework<'_>) {
        #[cfg(feature = "gui")]
        let gui_granted = fw.granted(vybe_bytecode::capabilities::Capability::Gui);

        if let Some(vm) = fw.vm.as_deref_mut() {
            // Always-on: the vybe:gui 2D drawing surface.
            crate::drawing::register(vm);

            // Widget-backed surface — only under `gui`, when Gui is granted and
            // this plugin carries the shared state.
            #[cfg(feature = "gui")]
            if gui_granted {
                if let Some(g) = &self.gui {
                    crate::gui::register(vm, g.clone());
                    crate::canvas::register(vm, g.clone());
                }
            }
        }
    }

    fn finalize(&self, fw: &mut vybe_bytecode::Framework<'_>) {
        // The vybe:gui control hierarchy + WinForms enums + control ctors,
        // registered via the `register_type` primitive after every plugin's
        // host fns exist. Idempotent across the base + gui-variant passes.
        crate::builtin_types::register_types(fw);
    }
}
