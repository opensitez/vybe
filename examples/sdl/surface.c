// examples/sdl/surface.c
//
// Doom's frame path in Doom's own shape: an offscreen 8-bit surface created
// with SDL_CreateRGBSurface, written through `surface->pixels`, then presented.
//
// This is the step past `paletted.c`, which passed a bare array. Here the
// surface is a real object with `w`/`h`/`pixels`/`format`, which is what
// `I_InitGraphics` builds and `R_DrawColumn` writes into.

#include <stdint.h>

typedef uint32_t Uint32;
typedef int32_t Sint32;
typedef void *SDL_Window;

typedef struct SDL_Palette { Uint32 *colors; } SDL_Palette;
typedef struct SDL_PixelFormat { SDL_Palette *palette; Sint32 BytesPerPixel; } SDL_PixelFormat;
typedef struct SDL_Surface {
    Sint32 w;
    Sint32 h;
    Sint32 depth;
    Sint32 pitch;
    unsigned char *pixels;
    SDL_PixelFormat *format;
} SDL_Surface;

#define SDL_INIT_VIDEO 0x00000020
#define SDL_WINDOW_SHOWN 0x00000004

extern int SDL_Init(Uint32 flags);
extern void SDL_Quit(void);
extern SDL_Window *SDL_CreateWindow(const char *t, Sint32 x, Sint32 y,
                                    Sint32 w, Sint32 h, Uint32 f);
extern void *SDL_GetWindowSurface(SDL_Window *window);
extern int SDL_ShowWindow(SDL_Window *window);
extern int SDL_UpdateWindowSurface(SDL_Window *window);
extern SDL_Surface *SDL_CreateRGBSurface(Uint32 flags, Sint32 w, Sint32 h,
                                         Sint32 depth, Uint32 rm, Uint32 gm,
                                         Uint32 bm, Uint32 am);
extern int SDL_BlitPaletted(void *dst, unsigned char *pixels, Sint32 w,
                            Sint32 h, Uint32 *palette, Sint32 dw, Sint32 dh);

#define SCREEN_W 320
#define SCREEN_H 200
#define WIN_W 960
#define WIN_H 600

static Uint32 palette[256];

int main(void) {
    if (SDL_Init(SDL_INIT_VIDEO) != 0) {
        return 1;
    }
    SDL_Window *window = SDL_CreateWindow("Vybe SDL - Surface Object",
                                          100, 100, WIN_W, WIN_H,
                                          SDL_WINDOW_SHOWN);
    void *winsurf = SDL_GetWindowSurface(window);
    SDL_ShowWindow(window);

    // The offscreen 8-bit buffer — `I_InitGraphics`'s `screenbuffer`.
    SDL_Surface *screen = SDL_CreateRGBSurface(0, SCREEN_W, SCREEN_H, 8,
                                               0, 0, 0, 0);
    if (screen == (void *)0) {
        SDL_Quit();
        return 1;
    }

    // A cool-toned ramp so this frame is distinguishable at a glance from
    // paletted.c's fire palette.
    for (Sint32 i = 0; i < 256; i = i + 1) {
        Sint32 b = 40 + i * 3;
        if (b > 255) { b = 255; }
        Sint32 g = i;
        Sint32 r = (i - 128) * 2;
        if (r < 0) { r = 0; }
        if (r > 255) { r = 255; }
        palette[i] = (Uint32)((r << 16) | (g << 8) | b);
    }

    // Write THROUGH the surface — this is the access pattern that matters:
    // `screen->pixels` and `screen->w`, exactly as Doom's renderer does it.
    unsigned char *px = screen->pixels;
    for (Sint32 y = 0; y < screen->h; y = y + 1) {
        for (Sint32 x = 0; x < screen->w; x = x + 1) {
            Sint32 v = (x ^ y) & 0xFF;            // XOR field — classic, and
            px[y * screen->w + x] = (unsigned char)v;  // every pixel distinct
        }
    }

    // Diagonal band, so a wrong pitch or row stride is obvious.
    for (Sint32 y = 0; y < screen->h; y = y + 1) {
        Sint32 x = (y * 3) % screen->w;
        px[y * screen->w + x] = 255;
    }

    SDL_BlitPaletted(winsurf, screen->pixels, screen->w, screen->h,
                     palette, WIN_W, WIN_H);
    SDL_UpdateWindowSurface(window);
    SDL_Quit();
    return 0;
}
