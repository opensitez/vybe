// examples/sdl/paletted.c
//
// The frame path a software renderer needs, in the shape Doom uses:
// an 8-bit palette-indexed buffer plus a 256-entry palette, expanded to RGBA
// and presented once per frame. This is `I_FinishUpdate` reduced to its
// essentials — see `sdlplan.md`.
//
// Renders a plasma field at Doom's native 320x200 and upscales it to the
// window, so it also shows whether scaling stays crisp (nearest-neighbour)
// rather than blurred (bilinear).

#include <stdint.h>

typedef uint32_t Uint32;
typedef int32_t Sint32;
typedef void *SDL_Window;
typedef void *SDL_Surface;

#define SDL_INIT_VIDEO 0x00000020
#define SDL_WINDOW_SHOWN 0x00000004

extern int SDL_Init(Uint32 flags);
extern void SDL_Quit(void);
extern SDL_Window *SDL_CreateWindow(const char *title, Sint32 x, Sint32 y,
                                    Sint32 w, Sint32 h, Uint32 flags);
extern SDL_Surface *SDL_GetWindowSurface(SDL_Window *window);
extern int SDL_ShowWindow(SDL_Window *window);
extern int SDL_UpdateWindowSurface(SDL_Window *window);

// Vybe frame path: pixels are palette indices, palette entries are 0xRRGGBB.
extern int SDL_BlitPaletted(SDL_Surface *surface, unsigned char *pixels,
                            Sint32 w, Sint32 h, Uint32 *palette,
                            Sint32 dst_w, Sint32 dst_h);

#define SCREEN_W 320
#define SCREEN_H 200
#define WIN_W 960
#define WIN_H 600

static unsigned char screen[SCREEN_W * SCREEN_H];
static Uint32 palette[256];

int main(void) {
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
    SDL_Surface *surface = SDL_GetWindowSurface(window);
    SDL_ShowWindow(window);

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
        palette[i] = (Uint32)((r << 16) | (g << 8) | b);
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
            screen[y * SCREEN_W + x] = (unsigned char)((d1 + d2) & 0xFF);
        }
    }

    // A hard-edged checker in the corner: with nearest-neighbour upscaling the
    // squares stay sharp; bilinear filtering visibly smears them.
    for (Sint32 y = 0; y < 32; y = y + 1) {
        for (Sint32 x = 0; x < 32; x = x + 1) {
            unsigned char v = (unsigned char)(((x / 4) + (y / 4)) % 2 ? 255 : 0);
            screen[y * SCREEN_W + x] = v;
        }
    }

    SDL_BlitPaletted(surface, screen, SCREEN_W, SCREEN_H, palette, WIN_W, WIN_H);
    SDL_UpdateWindowSurface(window);
    SDL_Quit();
    return 0;
}
