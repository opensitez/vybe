// A small software-rendered SDL frame in Doom's column-drawing shape.
// Compile with: cc doom_columns.c $(sdl2-config --cflags --libs) -o doom_columns
#include <SDL.h>
#include <assert.h>
#include <string.h>

#define WIDTH 320
#define HEIGHT 200
#define FRACBITS 16

static Uint8 *ylookup[HEIGHT];
static Uint8 source[128];
static Uint8 colormap[256];

static void draw_column(int pitch, int x, int frame)
{
    int yl = 20 + ((x + frame * 3) % 30);
    int yh = HEIGHT - 20 - ((x / 3 + frame) % 30);
    int frac = (frame * 7) << FRACBITS;
    int fracstep = (1 << FRACBITS) + ((x & 3) << 14);
    Uint8 *dest = ylookup[yl] + x;
    for (int y = yl; y <= yh; ++y) {
        *dest = colormap[source[(frac >> FRACBITS) & 127]];
        dest += pitch;
        frac += fracstep;
    }
}

int main(int argc, char **argv)
{
    int once = argc > 1 && strcmp(argv[1], "--once") == 0;
    assert(SDL_Init(SDL_INIT_VIDEO) == 0);
    SDL_Window *window = SDL_CreateWindow("Doom columns", 100, 100,
                                           WIDTH * 2, HEIGHT * 2, SDL_WINDOW_SHOWN);
    assert(window != NULL);
    SDL_Renderer *renderer = SDL_CreateRenderer(window, -1, SDL_RENDERER_SOFTWARE);
    assert(renderer != NULL);
    SDL_Surface *indexed = SDL_CreateRGBSurface(0, WIDTH, HEIGHT, 8, 0, 0, 0, 0);
    assert(indexed != NULL);
    assert(indexed->pitch == WIDTH);
    SDL_Surface *rgba = SDL_CreateRGBSurfaceWithFormatFrom(
        NULL, WIDTH, HEIGHT, 32, 0, SDL_PIXELFORMAT_ARGB8888);
    assert(rgba != NULL);
    SDL_Texture *texture = SDL_CreateTexture(renderer, SDL_PIXELFORMAT_ARGB8888,
                                             SDL_TEXTUREACCESS_STREAMING, WIDTH, HEIGHT);
    assert(texture != NULL);

    SDL_Color palette[256];
    for (int i = 0; i < 256; ++i) {
        palette[i].r = (Uint8)i;
        palette[i].g = (Uint8)((i * 3) & 255);
        palette[i].b = (Uint8)(255 - i);
        palette[i].a = 255;
        colormap[i] = (Uint8)i;
    }
    assert(SDL_SetPaletteColors(indexed->format->palette, palette, 0, 256) == 0);
    for (int i = 0; i < 128; ++i) source[i] = (Uint8)(i * 2);
    Uint8 *screen = (Uint8 *)indexed->pixels;
    for (int y = 0; y < HEIGHT; ++y) ylookup[y] = screen + y * indexed->pitch;
    assert(source[70] == 140);
    assert(colormap[140] == 140);
    assert(ylookup[100] == screen + 100 * indexed->pitch);

    SDL_Rect rect = {0, 0, WIDTH, HEIGHT};
    int running = 1;
    for (int frame = 0; running; ++frame) {
        memset(screen, 0, (size_t)indexed->pitch * HEIGHT);
        for (int x = 0; x < WIDTH; ++x)
            draw_column(indexed->pitch, x, frame);
        if (frame == 0) assert(screen[100 * indexed->pitch + 100] == 140);
        assert(SDL_LockTexture(texture, NULL, &rgba->pixels, &rgba->pitch) == 0);
        assert(SDL_LowerBlit(indexed, &rect, rgba, &rect) == 0);
        SDL_UnlockTexture(texture);
        assert(SDL_RenderCopy(renderer, texture, NULL, NULL) == 0);
        SDL_RenderPresent(renderer);
        SDL_Event event;
        while (SDL_PollEvent(&event)) {
            if (event.type == SDL_QUIT) running = 0;
        }
        if (once) running = 0;
        SDL_Delay(16);
    }

    SDL_DestroyTexture(texture);
    SDL_FreeSurface(rgba);
    SDL_FreeSurface(indexed);
    SDL_DestroyRenderer(renderer);
    SDL_DestroyWindow(window);
    SDL_Quit();
    return 0;
}
