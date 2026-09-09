//! `libc.*` namespace-tree registration for platform-owned libc surfaces.
//!
//! C also contributes profile-shaped libc entries, but platform math helpers
//! live here so any language can resolve `libc.math.*` without depending on
//! the C frontend being registered.

use std::sync::Once;

use vybe_compiler::primitives::namespaces::{self, NamespaceNode, Subtree};

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
            ("SDL_QuitSubSystem", "libc.sdl.SDL_QuitSubSystem"),
            ("SDL_GetError", "libc.sdl.SDL_GetError"),
            ("SDL_SetHint", "libc.sdl.SDL_SetHint"),
            (
                "SDL_SetHintWithPriority",
                "libc.sdl.SDL_SetHintWithPriority",
            ),
            ("SDL_FreeSurface", "libc.sdl.SDL_FreeSurface"),
            ("SDL_CreateRenderer", "libc.sdl.SDL_CreateRenderer"),
            ("SDL_DestroyRenderer", "libc.sdl.SDL_DestroyRenderer"),
            ("SDL_CreateTexture", "libc.sdl.SDL_CreateTexture"),
            ("SDL_DestroyTexture", "libc.sdl.SDL_DestroyTexture"),
            (
                "SDL_CreateRGBSurfaceFrom",
                "libc.sdl.SDL_CreateRGBSurfaceFrom",
            ),
            (
                "SDL_CreateRGBSurfaceWithFormatFrom",
                "libc.sdl.SDL_CreateRGBSurfaceWithFormatFrom",
            ),
            ("SDL_SetPaletteColors", "libc.sdl.SDL_SetPaletteColors"),
            ("SDL_LockTexture", "libc.sdl.SDL_LockTexture"),
            ("SDL_UnlockTexture", "libc.sdl.SDL_UnlockTexture"),
            ("SDL_LowerBlit", "libc.sdl.SDL_LowerBlit"),
            ("SDL_BlitSurface", "libc.sdl.SDL_BlitSurface"),
            ("SDL_RenderClear", "libc.sdl.SDL_RenderClear"),
            ("SDL_RenderCopy", "libc.sdl.SDL_RenderCopy"),
            ("SDL_RenderPresent", "libc.sdl.SDL_RenderPresent"),
            ("SDL_SetRenderTarget", "libc.sdl.SDL_SetRenderTarget"),
            ("SDL_SetRenderDrawColor", "libc.sdl.SDL_SetRenderDrawColor"),
            (
                "SDL_RenderSetLogicalSize",
                "libc.sdl.SDL_RenderSetLogicalSize",
            ),
            (
                "SDL_RenderSetIntegerScale",
                "libc.sdl.SDL_RenderSetIntegerScale",
            ),
            (
                "SDL_GetRendererOutputSize",
                "libc.sdl.SDL_GetRendererOutputSize",
            ),
            ("SDL_GetRendererInfo", "libc.sdl.SDL_GetRendererInfo"),
            ("SDL_GetWindowSize", "libc.sdl.SDL_GetWindowSize"),
            ("SDL_SetWindowSize", "libc.sdl.SDL_SetWindowSize"),
            ("SDL_SetWindowTitle", "libc.sdl.SDL_SetWindowTitle"),
            (
                "SDL_SetWindowFullscreen",
                "libc.sdl.SDL_SetWindowFullscreen",
            ),
            (
                "SDL_SetWindowMinimumSize",
                "libc.sdl.SDL_SetWindowMinimumSize",
            ),
            ("SDL_SetWindowIcon", "libc.sdl.SDL_SetWindowIcon"),
            ("SDL_GetWindowID", "libc.sdl.SDL_GetWindowID"),
            ("SDL_GetWindowFlags", "libc.sdl.SDL_GetWindowFlags"),
            (
                "SDL_GetWindowDisplayIndex",
                "libc.sdl.SDL_GetWindowDisplayIndex",
            ),
            (
                "SDL_GetNumVideoDisplays",
                "libc.sdl.SDL_GetNumVideoDisplays",
            ),
            ("SDL_GetDisplayBounds", "libc.sdl.SDL_GetDisplayBounds"),
            (
                "SDL_GetCurrentDisplayMode",
                "libc.sdl.SDL_GetCurrentDisplayMode",
            ),
            ("SDL_GetAppState", "libc.sdl.SDL_GetAppState"),
            (
                "SDL_SetRelativeMouseMode",
                "libc.sdl.SDL_SetRelativeMouseMode",
            ),
            (
                "SDL_GetRelativeMouseState",
                "libc.sdl.SDL_GetRelativeMouseState",
            ),
            ("SDL_WarpMouseInWindow", "libc.sdl.SDL_WarpMouseInWindow"),
            ("SDL_StartTextInput", "libc.sdl.SDL_StartTextInput"),
            ("SDL_StopTextInput", "libc.sdl.SDL_StopTextInput"),
            ("SDL_SetTextInputRect", "libc.sdl.SDL_SetTextInputRect"),
            ("SDL_IsTextInputActive", "libc.sdl.SDL_IsTextInputActive"),
            ("SDL_GetKeyFromScancode", "libc.sdl.SDL_GetKeyFromScancode"),
            ("SDL_GetVersion", "libc.sdl.SDL_GetVersion"),
            (
                "SDL_UpdateWindowSurfaceRects",
                "libc.sdl.SDL_UpdateWindowSurfaceRects",
            ),
            ("SDL_WaitEvent", "libc.sdl.SDL_WaitEvent"),
            (
                "SDL_CreateTextureFromSurface",
                "libc.sdl.SDL_CreateTextureFromSurface",
            ),
            ("SDL_SwapLE16", "libc.sdl.SDL_SwapLE16"),
            ("SDL_SwapLE32", "libc.sdl.SDL_SwapLE32"),
            ("SDL_SwapBE16", "libc.sdl.SDL_SwapBE16"),
            ("SDL_SwapBE32", "libc.sdl.SDL_SwapBE32"),
            ("SDL_max", "libc.sdl.SDL_max"),
            ("SDL_LockAudio", "libc.sdl.SDL_LockAudio"),
            ("SDL_UnlockAudio", "libc.sdl.SDL_UnlockAudio"),
            ("SDL_PauseAudio", "libc.sdl.SDL_PauseAudio"),
            ("SDL_BuildAudioCVT", "libc.sdl.SDL_BuildAudioCVT"),
            ("SDL_ConvertAudio", "libc.sdl.SDL_ConvertAudio"),
            ("SDL_MixAudioFormat", "libc.sdl.SDL_MixAudioFormat"),
            ("SDL_CreateMutex", "libc.sdl.SDL_CreateMutex"),
            ("SDL_DestroyMutex", "libc.sdl.SDL_DestroyMutex"),
            ("SDL_LockMutex", "libc.sdl.SDL_LockMutex"),
            ("SDL_UnlockMutex", "libc.sdl.SDL_UnlockMutex"),
            ("SDL_CreateCond", "libc.sdl.SDL_CreateCond"),
            ("SDL_DestroyCond", "libc.sdl.SDL_DestroyCond"),
            ("SDL_CondWait", "libc.sdl.SDL_CondWait"),
            ("SDL_CondSignal", "libc.sdl.SDL_CondSignal"),
            ("SDL_CreateThread", "libc.sdl.SDL_CreateThread"),
            ("SDL_WaitThread", "libc.sdl.SDL_WaitThread"),
            ("SDL_NumJoysticks", "libc.sdl.SDL_NumJoysticks"),
            ("SDL_IsGameController", "libc.sdl.SDL_IsGameController"),
            ("SDL_JoystickEventState", "libc.sdl.SDL_JoystickEventState"),
            ("SDL_JoystickOpen", "libc.sdl.SDL_JoystickOpen"),
            ("SDL_JoystickClose", "libc.sdl.SDL_JoystickClose"),
            ("SDL_JoystickName", "libc.sdl.SDL_JoystickName"),
            (
                "SDL_JoystickNameForIndex",
                "libc.sdl.SDL_JoystickNameForIndex",
            ),
            ("SDL_JoystickNumAxes", "libc.sdl.SDL_JoystickNumAxes"),
            ("SDL_JoystickNumButtons", "libc.sdl.SDL_JoystickNumButtons"),
            ("SDL_JoystickNumHats", "libc.sdl.SDL_JoystickNumHats"),
            ("SDL_JoystickGetAxis", "libc.sdl.SDL_JoystickGetAxis"),
            ("SDL_JoystickGetButton", "libc.sdl.SDL_JoystickGetButton"),
            ("SDL_JoystickGetHat", "libc.sdl.SDL_JoystickGetHat"),
            ("SDL_JoystickGetGUID", "libc.sdl.SDL_JoystickGetGUID"),
            (
                "SDL_JoystickGetDeviceGUID",
                "libc.sdl.SDL_JoystickGetDeviceGUID",
            ),
            (
                "SDL_JoystickGetGUIDString",
                "libc.sdl.SDL_JoystickGetGUIDString",
            ),
            (
                "SDL_JoystickGetGUIDFromString",
                "libc.sdl.SDL_JoystickGetGUIDFromString",
            ),
            (
                "SDL_GameControllerEventState",
                "libc.sdl.SDL_GameControllerEventState",
            ),
            ("SDL_GameControllerOpen", "libc.sdl.SDL_GameControllerOpen"),
            (
                "SDL_GameControllerClose",
                "libc.sdl.SDL_GameControllerClose",
            ),
            ("SDL_GameControllerName", "libc.sdl.SDL_GameControllerName"),
            ("SDL_GameControllerType", "libc.sdl.SDL_GameControllerType"),
            (
                "SDL_GameControllerTypeForIndex",
                "libc.sdl.SDL_GameControllerTypeForIndex",
            ),
            (
                "SDL_GameControllerGetAxis",
                "libc.sdl.SDL_GameControllerGetAxis",
            ),
            (
                "SDL_GameControllerGetButton",
                "libc.sdl.SDL_GameControllerGetButton",
            ),
            (
                "SDL_GameControllerMappingForGUID",
                "libc.sdl.SDL_GameControllerMappingForGUID",
            ),
            ("Mix_Init", "libc.sdl.Mix_Init"),
            ("Mix_OpenAudioDevice", "libc.sdl.Mix_OpenAudioDevice"),
            ("Mix_CloseAudio", "libc.sdl.Mix_CloseAudio"),
            ("Mix_QuerySpec", "libc.sdl.Mix_QuerySpec"),
            ("Mix_GetError", "libc.sdl.Mix_GetError"),
            ("Mix_LoadMUS", "libc.sdl.Mix_LoadMUS"),
            ("Mix_LoadMUS_RW", "libc.sdl.Mix_LoadMUS_RW"),
            ("Mix_FreeMusic", "libc.sdl.Mix_FreeMusic"),
            ("Mix_PlayMusic", "libc.sdl.Mix_PlayMusic"),
            ("Mix_HaltMusic", "libc.sdl.Mix_HaltMusic"),
            ("Mix_PauseMusic", "libc.sdl.Mix_PauseMusic"),
            ("Mix_ResumeMusic", "libc.sdl.Mix_ResumeMusic"),
            ("Mix_PlayingMusic", "libc.sdl.Mix_PlayingMusic"),
            ("Mix_Playing", "libc.sdl.Mix_Playing"),
            ("Mix_VolumeMusic", "libc.sdl.Mix_VolumeMusic"),
            ("Mix_SetMusicPosition", "libc.sdl.Mix_SetMusicPosition"),
            ("Mix_SetMusicCMD", "libc.sdl.Mix_SetMusicCMD"),
            ("Mix_HookMusic", "libc.sdl.Mix_HookMusic"),
            ("Mix_RegisterEffect", "libc.sdl.Mix_RegisterEffect"),
            ("Mix_AllocateChannels", "libc.sdl.Mix_AllocateChannels"),
            ("Mix_PlayChannel", "libc.sdl.Mix_PlayChannel"),
            ("Mix_HaltChannel", "libc.sdl.Mix_HaltChannel"),
            ("Mix_SetPanning", "libc.sdl.Mix_SetPanning"),
            ("SDLNet_Init", "libc.sdl.SDLNet_Init"),
            ("SDLNet_GetError", "libc.sdl.SDLNet_GetError"),
            ("SDLNet_UDP_Open", "libc.sdl.SDLNet_UDP_Open"),
            ("SDLNet_AllocPacket", "libc.sdl.SDLNet_AllocPacket"),
            ("SDLNet_ResolveHost", "libc.sdl.SDLNet_ResolveHost"),
            ("SDLNet_UDP_Send", "libc.sdl.SDLNet_UDP_Send"),
            ("SDLNet_UDP_Recv", "libc.sdl.SDLNet_UDP_Recv"),
            ("SDLNet_Read16", "libc.sdl.SDLNet_Read16"),
            ("SDLNet_Read32", "libc.sdl.SDLNet_Read32"),
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
        namespaces::register_namespace_tree("sqlite3", NamespaceNode::Namespace(sqlite.clone()));
        namespaces::register_namespace_tree("mysql", NamespaceNode::Namespace(mysql.clone()));

        // ⛔ THE LEAVES, not just the package node. `sqlite3_open` is a FREE
        // FUNCTION in C — the source never writes `sqlite3.sqlite3_open` — so
        // resolution lands on a bare key and `resolve_key` looks it up at the
        // ROOT of the tree. Registering only the `sqlite3` namespace put every
        // symbol one level too deep: the package resolved, the call did not,
        // and both C and Fortran got `undefined is not callable`.
        //
        // namespaceplan.md: "Leaves: every callable/readable endpoint that
        // resolution may land on. A package or type node alone is incomplete."
        // Registered here rather than as `[builtins]` rows in each consuming
        // profile, so one registration serves C, Fortran and anything else that
        // names the C ABI.
        for (symbol, node) in sqlite.iter().chain(mysql.iter()) {
            namespaces::register_namespace_tree(symbol, node.clone());
        }

        // Keep compatibility with existing C profile entries that emit
        // `common:sdl.*` without a libc prefix.
        namespaces::register_namespace_tree("sdl", NamespaceNode::Namespace(sdl));
    });
}
