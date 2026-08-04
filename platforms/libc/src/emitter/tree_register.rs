//! `libc.*` namespace-tree registration for platform-owned libc surfaces.
//!
//! C also contributes profile-shaped libc entries, but platform math helpers
//! live here so any language can resolve `libc.math.*` without depending on
//! the C frontend being registered.

use std::sync::Once;

use vybe_runtime::namespaces::{self, NamespaceNode, Subtree};

/// Register the platform libc surface under the `libc` root. Idempotent;
/// later C/profile registration merges with this tree.
pub fn register_namespace_tree() {
    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        let mut math = Subtree::new();
        let mut sdl = Subtree::new();
        for (name, emit) in [
            ("erf", "libc.math.erf"),
            ("erfc", "libc.math.erfc"),
            ("tgamma", "libc.math.tgamma"),
            ("gamma", "libc.math.tgamma"),
            ("lgamma", "libc.math.lgamma"),
        ] {
            math.insert(
                name.to_string(),
                NamespaceNode::CommonEmit(emit.to_string()),
            );
        }

        for (name, emit) in [
            ("SDL_Init", "libc.sdl.SDL_Init"),
            ("SDL_InitSubSystem", "libc.sdl.SDL_InitSubSystem"),
            ("SDL_Quit", "libc.sdl.SDL_Quit"),
            ("SDL_CreateWindow", "libc.sdl.SDL_CreateWindow"),
            ("SDL_DestroyWindow", "libc.sdl.SDL_DestroyWindow"),
            ("SDL_GetWindowSurface", "libc.sdl.SDL_GetWindowSurface"),
            ("SDL_FillRect", "libc.sdl.SDL_FillRect"),
            ("SDL_UpdateWindowSurface", "libc.sdl.SDL_UpdateWindowSurface"),
            ("SDL_Delay", "libc.sdl.SDL_Delay"),
            ("SDL_MapRGB", "libc.sdl.SDL_MapRGB"),
            ("SDL_MapRGBA", "libc.sdl.SDL_MapRGBA"),
            ("SDL_DrawText", "libc.sdl.SDL_DrawText"),
            ("SDL_DrawLine", "libc.sdl.SDL_DrawLine"),
            ("SDL_ShowWindow", "libc.sdl.SDL_ShowWindow"),
            ("SDL_HideWindow", "libc.sdl.SDL_HideWindow"),
            ("SDL_BlitPaletted", "libc.sdl.SDL_BlitPaletted"),
            ("SDL_ShowSimpleMessageBox", "libc.sdl.SDL_ShowSimpleMessageBox"),
        ] {
            sdl.insert(name.to_string(), NamespaceNode::CommonEmit(emit.to_string()));
        }

        let mut root = Subtree::new();
        root.insert("math".to_string(), NamespaceNode::Namespace(math));
        root.insert("sdl".to_string(), NamespaceNode::Namespace(sdl.clone()));
        namespaces::register_namespace_tree("libc", NamespaceNode::Namespace(root));

        // Keep compatibility with existing C profile entries that emit
        // `common:sdl.*` without a libc prefix.
        namespaces::register_namespace_tree("sdl", NamespaceNode::Namespace(sdl));
    });
}
