// examples/sdl/paletted.c
//
// The frame path a software renderer needs, in the shape Doom uses:
// an 8-bit surface and SDL_Color palette, converted into a locked ARGB8888
// texture by SDL_LowerBlit, then presented by SDL_RenderCopy.
// Uses only standard SDL2 APIs. Pass --once for a single-frame smoke test.
//
// Renders a plasma field at Doom's native 320x200 and upscales it to the
// window, so it also shows whether scaling stays crisp (nearest-neighbour)
// rather than blurred (bilinear).

#include <assert.h>
#include <string.h>
#include <SDL.h>

#define SCREEN_W 320
#define SCREEN_H 200
#define WIN_W 960
#define WIN_H 600

static SDL_Color palette[256];

int main(int argc, char **argv) {
    if (SDL_Init(SDL_INIT_VIDEO) != 0) {
        return 1;
    }
    SDL_Window *window = SDL_CreateWindow("Vybe SDL - Paletted Frame",
                                          100, 100, WIN_W, WIN_H,
                                          SDL_WINDOW_SHOWN);
    if (window == (void *)0) {
        SDL_Quit();
        return 1;
    }
    SDL_SetHint(SDL_HINT_RENDER_SCALE_QUALITY, "nearest");
    SDL_Renderer *renderer = SDL_CreateRenderer(window, -1, SDL_RENDERER_SOFTWARE);
    assert(renderer != NULL);
    SDL_Surface *surface = SDL_CreateRGBSurface(0, SCREEN_W, SCREEN_H, 8, 0, 0, 0, 0);
    assert(surface != NULL);
    SDL_Surface *argb = SDL_CreateRGBSurfaceWithFormatFrom(
        NULL, SCREEN_W, SCREEN_H, 32, 0, SDL_PIXELFORMAT_ARGB8888);
    assert(argb != NULL);
    SDL_Texture *texture = SDL_CreateTexture(renderer, SDL_PIXELFORMAT_ARGB8888,
        SDL_TEXTUREACCESS_STREAMING, SCREEN_W, SCREEN_H);
    assert(texture != NULL);
    unsigned char *screen = surface->pixels;
    int pitch = surface->pitch;

    // A fire-ish ramp: black → red → orange → yellow → white, the kind of
    // palette a software renderer actually ships.
    for (Sint32 i = 0; i < 256; i = i + 1) {
        Sint32 r = i * 3;
        if (r > 255) { r = 255; }
        Sint32 g = (i - 64) * 3;
        if (g < 0) { g = 0; }
        if (g > 255) { g = 255; }
        Sint32 b = (i - 176) * 4;
        if (b < 0) { b = 0; }
        if (b > 255) { b = 255; }
        palette[i].r = (Uint8) r;
        palette[i].g = (Uint8) g;
        palette[i].b = (Uint8) b;
        palette[i].a = 255;
    }

    // Plasma: integer-only, no math.h. Concentric interference from two
    // sources gives large flat regions AND fine detail, so both the palette
    // mapping and the scaling filter are easy to judge by eye.
    for (Sint32 y = 0; y < SCREEN_H; y = y + 1) {
        for (Sint32 x = 0; x < SCREEN_W; x = x + 1) {
            Sint32 dx = x - SCREEN_W / 2;
            Sint32 dy = y - SCREEN_H / 2;
            Sint32 d1 = (dx * dx + dy * dy) / 96;
            Sint32 ex = x - 40;
            Sint32 ey = y - 30;
            Sint32 d2 = (ex * ex + ey * ey) / 64;
            screen[y * pitch + x] = (unsigned char)((d1 + d2) & 0xFF);
        }
    }

    // A hard-edged checker in the corner: with nearest-neighbour upscaling the
    // squares stay sharp; bilinear filtering visibly smears them.
    for (Sint32 y = 0; y < 32; y = y + 1) {
        for (Sint32 x = 0; x < 32; x = x + 1) {
            unsigned char v = (unsigned char)(((x / 4) + (y / 4)) % 2 ? 255 : 0);
            screen[y * pitch + x] = v;
        }
    }

    assert(SDL_SetPaletteColors(surface->format->palette, palette, 0, 256) == 0);
    SDL_Rect rect = {0, 0, SCREEN_W, SCREEN_H};
    assert(SDL_LockTexture(texture, NULL, &argb->pixels, &argb->pitch) == 0);
    assert(SDL_LowerBlit(surface, &rect, argb, &rect) == 0);
    SDL_UnlockTexture(texture);
    assert(SDL_RenderCopy(renderer, texture, NULL, NULL) == 0);
    SDL_RenderPresent(renderer);

    int running = !(argc > 1 && strcmp(argv[1], "--once") == 0);
    while (running) {
        SDL_Event event;
        while (SDL_PollEvent(&event)) {
            if (event.type == SDL_QUIT) running = 0;
        }
        SDL_Delay(16);
    }
    SDL_DestroyTexture(texture);
    SDL_FreeSurface(argb);
    SDL_FreeSurface(surface);
    SDL_DestroyRenderer(renderer);
    SDL_DestroyWindow(window);
    SDL_Quit();
    return 0;
}
