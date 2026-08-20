/* Multi-file SDL calculator — the Doom-shaped integration example:
 * real headers (SDL2/SDL.h, math.h, stdio.h — all GATED includes),
 * mouse + keyboard input, an SDL surface, several translation units,
 * and a Makefile. `make test` drives the same event path headlessly
 * through SDL_PushEvent via the --entry override. */
#include <SDL2/SDL.h>
#include "calc.h"
#include "ui.h"

int main(void) {
    struct Calc c;
    SDL_Window *win;
    SDL_Surface *screen;
    SDL_Event e;
    int running = 1;

    SDL_Init(SDL_INIT_VIDEO);
    win = SDL_CreateWindow("vybe calc", SDL_WINDOWPOS_CENTERED,
                           SDL_WINDOWPOS_CENTERED, CALC_W, CALC_H, 0);
    screen = SDL_GetWindowSurface(win);
    calc_reset(&c);

    while (running) {
        while (SDL_PollEvent(&e)) {
            running = ui_handle_event(&c, &e);
        }
        ui_render(screen, &c);
        SDL_UpdateWindowSurface(win);
        SDL_Delay(16);
    }
    SDL_Quit();
    return 0;
}
