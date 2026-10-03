/* Small SDL input and redraw check. Click to move the square; arrow keys move
 * it by one step. Escape or closing the window exits. Build natively with:
 * cc sdl_input_game.c $(pkg-config --cflags --libs sdl2) -o sdl_input_game
 */
#include <SDL2/SDL.h>
#include <stdio.h>

#define WIDTH 320
#define HEIGHT 200
#define SIZE 16

static int x = (WIDTH - SIZE) / 2;
static int y = (HEIGHT - SIZE) / 2;

static void draw(SDL_Window *window, SDL_Surface *surface) {
    SDL_Rect square = { x, y, SIZE, SIZE };
    SDL_FillRect(surface, NULL, SDL_MapRGB(surface->format, 18, 28, 38));
    SDL_FillRect(surface, &square, SDL_MapRGB(surface->format, 255, 200, 60));
    SDL_UpdateWindowSurface(window);
}

static int handle_event(const SDL_Event *event) {
    if (event->type == SDL_QUIT) return 0;
    if (event->type == SDL_MOUSEBUTTONDOWN) {
        printf("mouse button=%d x=%d y=%d\n", event->button.button,
               event->button.x, event->button.y);
        x = event->button.x - SIZE / 2;
        y = event->button.y - SIZE / 2;
    } else if (event->type == SDL_KEYDOWN) {
        printf("key sym=%d scancode=%d\n", event->key.keysym.sym,
               event->key.keysym.scancode);
        switch (event->key.keysym.sym) {
        case SDLK_ESCAPE: return 0;
        case SDLK_LEFT: x -= SIZE; break;
        case SDLK_RIGHT: x += SIZE; break;
        case SDLK_UP: y -= SIZE; break;
        case SDLK_DOWN: y += SIZE; break;
        }
    }
    if (x < 0) x = 0;
    if (y < 0) y = 0;
    if (x > WIDTH - SIZE) x = WIDTH - SIZE;
    if (y > HEIGHT - SIZE) y = HEIGHT - SIZE;
    return 1;
}

int main(void) {
    SDL_Window *window;
    SDL_Surface *surface;
    SDL_Event event;
    int running = 1;
    if (SDL_Init(SDL_INIT_VIDEO) != 0) return 1;
    window = SDL_CreateWindow("SDL input check", SDL_WINDOWPOS_CENTERED,
                              SDL_WINDOWPOS_CENTERED, WIDTH, HEIGHT, 0);
    if (window == NULL) return 2;
    surface = SDL_GetWindowSurface(window);
    if (surface == NULL) return 3;
    draw(window, surface);
    while (running) {
        while (SDL_PollEvent(&event)) {
            running = handle_event(&event);
            if (!running) break;
            draw(window, surface);
        }
        SDL_Delay(16);
    }
    SDL_DestroyWindow(window);
    SDL_Quit();
    return 0;
}
