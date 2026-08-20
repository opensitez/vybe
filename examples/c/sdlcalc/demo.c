/* One-frame demo: scripted input, one render, then RETURN — the window
 * presents the final frame after exit (today's presentation model).
 * Interactive play needs per-frame presentation (JSPI) — the last
 * architecture item before Doom. Run: make demo   or   make shot */
#include <SDL2/SDL.h>
#include "calc.h"
#include "ui.h"
#include <stdio.h>

static void push_key(int sym) {
    SDL_Event e;
    e.type = SDL_KEYDOWN;
    e.key.keysym.sym = sym;
    SDL_PushEvent(&e);
}

int demo(void) {
    struct Calc c;
    SDL_Window *win;
    SDL_Surface *screen;
    SDL_Event e;

    SDL_Init(SDL_INIT_VIDEO);
    win = SDL_CreateWindow("vybe calc", SDL_WINDOWPOS_CENTERED,
                           SDL_WINDOWPOS_CENTERED, CALC_W, CALC_H, 0);
    screen = SDL_GetWindowSurface(win);
    calc_reset(&c);

    /* scripted: 1.5 + 2.25 =  → display shows 3.75 */
    push_key(SDLK_1); push_key(SDLK_PERIOD); push_key(SDLK_5);
    push_key(SDLK_PLUS);
    push_key(SDLK_2); push_key(SDLK_PERIOD); push_key(SDLK_2); push_key(SDLK_5);
    push_key(SDLK_RETURN);
    while (SDL_PollEvent(&e)) { ui_handle_event(&c, &e); }

    printf("display would read: %g\n", calc_value(&c));
    ui_render(screen, &c);
    SDL_UpdateWindowSurface(win);
    return 0;
}
int main(void) { return demo(); }
