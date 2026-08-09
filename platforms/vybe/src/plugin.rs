//! The vybe platform as a `vybe_runtime::Plugin` — one plugin, same type as all
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
#[derive(Default)]
pub struct Plugin;

/// The widget state this plugin registers over. It lives BESIDE the plugin
/// rather than inside the value because the registry holds `&'static dyn
/// Plugin` — a plugin is identified by what it registers, not by an instance
/// someone constructed. This is the shape the doc above anticipated for a
/// dylib: the plugin's own `init` finds the state, and the accessor hands back
/// a registered handle.
#[cfg(feature = "gui")]
static GUI: std::sync::Mutex<Option<std::sync::Arc<std::sync::Mutex<crate::gui_state::GuiState>>>> =
    std::sync::Mutex::new(None);

impl Plugin {
    /// A drawing-only vybe plugin (no widget state).
    pub fn new() -> Self {
        Plugin
    }

    /// Install a freshly created `GuiState` and return the shared handle. Each
    /// call replaces the previous one, so a new VM starts from clean widget
    /// state — the behaviour of the old per-instance `with_gui()`.
    #[cfg(feature = "gui")]
    pub fn with_gui() -> Self {
        let state = std::sync::Arc::new(std::sync::Mutex::new(crate::gui_state::GuiState::new()));
        *GUI.lock().unwrap() = Some(state);
        Plugin
    }

    /// The shared `GuiState`, if one is installed (for the form launcher).
    #[cfg(feature = "gui")]
    pub fn gui_state(
        &self,
    ) -> Option<std::sync::Arc<std::sync::Mutex<crate::gui_state::GuiState>>> {
        GUI.lock().unwrap().clone()
    }
}

impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "vybe"
    }

    fn init(&self, fw: &mut vybe_runtime::Framework<'_>) {
        // Whether the widget-backed surface got registered below. When it
        // did not, this plugin installs its own no-op `vybe:gui` stubs so
        // compiled control/form code still links. That fallback used to live
        // in the compiler, which had to name this crate to call it — a
        // plugin installs its own surface.
        #[allow(unused_mut)]
        let mut widgets_registered = false;
        #[cfg(feature = "gui")]
        let gui_granted = fw.granted(vybe_runtime::capabilities::Capability::Gui);

        if let Some(vm) = fw.vm.as_deref_mut() {
            // Always-on: the vybe:gui 2D drawing surface.
            crate::drawing::register(vm);

            // Widget-backed surface — only under `gui`, when Gui is granted and
            // this plugin carries the shared state.
            #[cfg(feature = "gui")]
            if gui_granted {
                if let Some(g) = GUI.lock().unwrap().as_ref() {
                    widgets_registered = true;
                    crate::gui::register(vm, g.clone());
                    crate::canvas_backend_impl::install(g.clone());
                }
            }

            // No widget surface — install the no-op stubs so compiled
            // control/form code still links.
            if !widgets_registered {
                crate::register_gui_stubs(vm);
            }
        }
    }

    fn finalize(&self, fw: &mut vybe_runtime::Framework<'_>) {
        // The vybe:gui control hierarchy + WinForms enums + control ctors,
        // registered via the `register_type` primitive after every plugin's
        // host fns exist. Idempotent across the base + gui-variant passes.
        crate::builtin_types::register_types(fw);
    }

    /// Drop the script's GUI: controls, event handlers, the property store,
    /// the form object, canvases.
    ///
    /// `GuiState::reset` was written for exactly this and its doc already
    /// claimed "the runner calls this ... as part of `reset_to`" — but it had
    /// no callers at all. Widget state therefore survived every reset, so a
    /// reused VM handed the next program the previous one's window, controls
    /// and handlers, with the handlers still holding guest closures.
    fn reset(&self) {
        #[cfg(feature = "gui")]
        if let Some(state) = GUI.lock().unwrap().as_ref() {
            if let Ok(mut g) = state.lock() {
                g.reset();
            }
        }
        // `GuiState` is this crate's view of the widgets; the widget crate
        // keeps its OWN process-wide tables, and those outlive a `GuiState`
        // replacement. Documents, undelivered input, and scheduled timers and
        // frames all belong to the program that produced them.
        #[cfg(feature = "gui")]
        {
            vybe_widgets::dom::reset();
            vybe_widgets::ui_events::reset();
            vybe_widgets::scheduling::reset();
        }
    }
}

/// Install a fresh widget `GuiState`, then run phase 1 of the ONE registration
/// loop over every linked plugin. Returns the shared handle for the form
/// launcher / test assertions. Pair with `finalize_platforms`.
///
/// This lives here, not in the compiler: the gui-variant needs this crate's
/// `GuiState`, and the compiler must not name a platform crate.
#[cfg(feature = "gui")]
pub fn init_platforms_with_gui(
    vm: &mut vybe_runtime::VM,
) -> std::sync::Arc<std::sync::Mutex<crate::gui_state::GuiState>> {
    let plugin = Plugin::with_gui();
    vybe_runtime::init_registered_plugins(vm, &vybe_runtime::capabilities::Capabilities::all());
    plugin
        .gui_state()
        .expect("with_gui() always installs a GuiState")
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);
