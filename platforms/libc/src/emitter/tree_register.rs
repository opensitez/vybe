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
            // libm declares all six in <math.h>; registering them here is what
            // makes them reachable from every language, not just C.
            ("j0", "libc.math.j0"),
            ("j1", "libc.math.j1"),
            ("y0", "libc.math.y0"),
            ("y1", "libc.math.y1"),
            ("jn", "libc.math.jn"),
            ("yn", "libc.math.yn"),
            ("erfcx", "libc.math.erfcx"),
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
            (
                "SDL_UpdateWindowSurface",
                "libc.sdl.SDL_UpdateWindowSurface",
            ),
            ("SDL_Delay", "libc.sdl.SDL_Delay"),
            ("SDL_MapRGB", "libc.sdl.SDL_MapRGB"),
            ("SDL_MapRGBA", "libc.sdl.SDL_MapRGBA"),
            ("SDL_DrawText", "libc.sdl.SDL_DrawText"),
            ("SDL_DrawLine", "libc.sdl.SDL_DrawLine"),
            ("SDL_ShowWindow", "libc.sdl.SDL_ShowWindow"),
            ("SDL_HideWindow", "libc.sdl.SDL_HideWindow"),
            ("SDL_BlitPaletted", "libc.sdl.SDL_BlitPaletted"),
            (
                "SDL_ShowSimpleMessageBox",
                "libc.sdl.SDL_ShowSimpleMessageBox",
            ),
        ] {
            sdl.insert(
                name.to_string(),
                NamespaceNode::CommonEmit(emit.to_string()),
            );
        }

        // C ABI database surfaces. Registered like `sdl`: under `libc.*` for a
        // qualified reach, and as a bare root so the plain C symbol resolves —
        // which is what a Fortran `bind(c, name="sqlite3_open")` emits.
        let mut sqlite = Subtree::new();
        for name in [
            "sqlite3_open",
            "sqlite3_open_v2",
            "sqlite3_close",
            "sqlite3_close_v2",
            "sqlite3_exec",
            "sqlite3_prepare",
            "sqlite3_prepare_v2",
            "sqlite3_finalize",
            "sqlite3_reset",
            "sqlite3_errmsg",
            "sqlite3_errcode",
            "sqlite3_extended_errcode",
        ] {
            sqlite.insert(
                name.to_string(),
                NamespaceNode::CommonEmit(format!("libc.sqlite3.{name}")),
            );
        }

        let mut mysql = Subtree::new();
        for name in [
            "mysql_init",
            "mysql_real_connect",
            "mysql_close",
            "mysql_select_db",
            "mysql_query",
            "mysql_real_query",
            "mysql_store_result",
            "mysql_use_result",
            "mysql_free_result",
        ] {
            mysql.insert(
                name.to_string(),
                NamespaceNode::CommonEmit(format!("libc.mysql.{name}")),
            );
        }

        let mut root = Subtree::new();
        root.insert("math".to_string(), NamespaceNode::Namespace(math));
        root.insert("sdl".to_string(), NamespaceNode::Namespace(sdl.clone()));
        root.insert(
            "sqlite3".to_string(),
            NamespaceNode::Namespace(sqlite.clone()),
        );
        root.insert("mysql".to_string(), NamespaceNode::Namespace(mysql.clone()));
        namespaces::register_namespace_tree("libc", NamespaceNode::Namespace(root));
        namespaces::register_namespace_tree("sqlite3", NamespaceNode::Namespace(sqlite));
        namespaces::register_namespace_tree("mysql", NamespaceNode::Namespace(mysql));

        // Keep compatibility with existing C profile entries that emit
        // `common:sdl.*` without a libc prefix.
        namespaces::register_namespace_tree("sdl", NamespaceNode::Namespace(sdl));
    });
}
