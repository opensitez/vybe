use dioxus::prelude::*;
use crate::app_state::AppState;
use vybe_ui::runtime_panel::RuntimeProject;

/// Editor wrapper – provides the project from AppState as a RuntimeProject
/// context, then delegates entirely to the shared FormRunner in vybe_ui.
#[component]
pub fn RuntimePanel() -> Element {
    let mut state = use_context::<AppState>();

    let finished = use_signal(|| false);

    // Bridge: expose the editor's project signal via the shared RuntimeProject
    // context so FormRunner can read it.
    use_context_provider(|| RuntimeProject {
        project: state.project,
        finished,
    });

    // When the FormRunner signals that a console project finished,
    // automatically leave run-mode.
    use_effect(move || {
        if *finished.read() {
            state.run_mode.set(false);
        }
    });

    rsx! {
        vybe_ui::FormRunner {}
    }
}
