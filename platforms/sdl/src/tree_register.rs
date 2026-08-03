//! `sdl` namespace registration.
//!
//! This adapter exposes SDL-shaped symbols and routes through
//! `platforms/sdl` emitter dispatch (`emit_*` in this crate), with runtime
//! behavior provided by existing `vybe:gui` host functions.

use std::sync::Once;

use vybe_runtime::namespaces::{self, NamespaceNode, Subtree};

fn common_emit(name: &str) -> NamespaceNode {
    NamespaceNode::CommonEmit(name.to_string())
}

fn insert_common(root: &mut Subtree, name: &str, op: &str) {
    root.insert(name.to_string(), common_emit(op));
}

/// Register the `sdl` namespace tree.
pub fn register_namespace_tree() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        let mut root = Subtree::new();

        let entries = [
            ("SDL_Init", "sdl.SDL_Init"),
            ("SDL_InitSubSystem", "sdl.SDL_InitSubSystem"),
            ("SDL_Quit", "sdl.SDL_Quit"),
            ("SDL_CreateWindow", "sdl.SDL_CreateWindow"),
            ("SDL_DestroyWindow", "sdl.SDL_DestroyWindow"),
            ("SDL_GetWindowSurface", "sdl.SDL_GetWindowSurface"),
            ("SDL_FillRect", "sdl.SDL_FillRect"),
            ("SDL_UpdateWindowSurface", "sdl.SDL_UpdateWindowSurface"),
            ("SDL_Delay", "sdl.SDL_Delay"),
            ("SDL_MapRGB", "sdl.SDL_MapRGB"),
            ("SDL_MapRGBA", "sdl.SDL_MapRGBA"),
            ("SDL_ShowWindow", "sdl.SDL_ShowWindow"),
            ("SDL_HideWindow", "sdl.SDL_HideWindow"),
            ("SDL_ShowSimpleMessageBox", "sdl.SDL_ShowSimpleMessageBox"),
        ];

        for (name, op) in entries {
            insert_common(&mut root, name, op);
        }

        namespaces::register_namespace_tree("sdl", NamespaceNode::Namespace(root));
    });
}
