#include <stdint.h>

typedef uint32_t Uint32;
typedef int32_t Sint32;
typedef void *SDL_Window;
typedef void *SDL_Renderer;
typedef void *SDL_Texture;

typedef struct SDL_Palette {
    Uint32 *colors;
} SDL_Palette;

typedef struct SDL_PixelFormat {
    SDL_Palette *palette;
    Sint32 BytesPerPixel;
} SDL_PixelFormat;

typedef struct SDL_Surface {
    Sint32 w;
    Sint32 h;
    Sint32 depth;
    Sint32 pitch;
    unsigned char *pixels;
    SDL_PixelFormat *format;
} SDL_Surface;

extern int SDL_Init(Uint32 flags);
extern void SDL_Quit(void);
extern SDL_Window *SDL_CreateWindow(const char *title, Sint32 x, Sint32 y,
                                    Sint32 w, Sint32 h, Uint32 flags);
extern SDL_Renderer *SDL_CreateRenderer(SDL_Window *window, Sint32 index,
                                        Uint32 flags);
extern SDL_Surface *SDL_CreateRGBSurface(Uint32 flags, Sint32 w, Sint32 h,
                                         Sint32 depth, Uint32 rmask,
                                         Uint32 gmask, Uint32 bmask,
                                         Uint32 amask);
extern SDL_Texture *SDL_CreateTextureFromSurface(SDL_Renderer *renderer,
                                                 SDL_Surface *surface);
extern int SDL_RenderCopy(SDL_Renderer *renderer, SDL_Texture *texture,
                          void *srcrect, void *dstrect);

static Uint32 palette[256];

int main(void) {
    SDL_Init(32);
    SDL_Window *window = SDL_CreateWindow("render-copy", 0, 0, 64, 64, 4);
    SDL_Renderer *renderer = SDL_CreateRenderer(window, -1, 0);
    SDL_Surface *surface = SDL_CreateRGBSurface(0, 8, 8, 8, 0, 0, 0, 0);

    for (Sint32 i = 0; i < 256; i = i + 1) {
        palette[i] = (Uint32)((i << 16) | (i << 8) | i);
    }
    surface->format->palette->colors = palette;

    unsigned char *pixels = surface->pixels;
    for (Sint32 i = 0; i < 64; i = i + 1) {
        pixels[i] = (unsigned char)i;
    }

    SDL_Texture *texture = SDL_CreateTextureFromSurface(renderer, surface);
    int rc = SDL_RenderCopy(renderer, texture, 0, 0);
    SDL_Quit();
    return rc;
}
