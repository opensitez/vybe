use vybe_runtime::Chunk;

pub fn dispatch(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    match name {
        "sdl.SDL_Init" => {
            super::sdl::emit_sdl_init(chunks, current, argc, line);
            true
        }
        "sdl.SDL_InitSubSystem" => {
            super::sdl::emit_sdl_init_subsystem(chunks, current, argc, line);
            true
        }
        "sdl.SDL_Quit" => {
            super::sdl::emit_sdl_quit(chunks, current, line);
            true
        }
        "sdl.SDL_CreateWindow" => {
            super::sdl::emit_sdl_create_window(chunks, current, argc, line);
            true
        }
        "sdl.SDL_DestroyWindow" => {
            super::sdl::emit_sdl_destroy_window(chunks, current, argc, line);
            true
        }
        "sdl.SDL_GetWindowSurface" => {
            super::sdl::emit_sdl_get_window_surface(chunks, current, argc, line);
            true
        }
        "sdl.SDL_FillRect" => {
            super::sdl::emit_sdl_fill_rect(chunks, current, argc, line);
            true
        }
        "sdl.SDL_UpdateWindowSurface" => {
            super::sdl::emit_sdl_update_window_surface(chunks, current, argc, line);
            true
        }
        "sdl.SDL_Delay" => {
            super::sdl::emit_sdl_delay(chunks, current, argc, line);
            true
        }
        "sdl.SDL_MapRGB" => {
            super::sdl::emit_sdl_map_rgb(chunks, current, argc, line);
            true
        }
        "sdl.SDL_MapRGBA" => {
            super::sdl::emit_sdl_map_rgba(chunks, current, argc, line);
            true
        }
        "sdl.SDL_ShowWindow" => {
            super::sdl::emit_sdl_show_window(chunks, current, argc, line);
            true
        }
        "sdl.SDL_HideWindow" => {
            super::sdl::emit_sdl_hide_window(chunks, current, argc, line);
            true
        }
        "sdl.SDL_ShowSimpleMessageBox" => {
            super::sdl::emit_sdl_show_simple_message_box(chunks, current, argc, line);
            true
        }
        _ => false,
    }
}
