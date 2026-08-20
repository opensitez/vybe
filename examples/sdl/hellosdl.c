// examples/sdl/hellosdl.c
//
// SDL-style sample for the Vybe SDL adapter (`platforms/libc/src/emitter/sdl.rs`).
//
// A small signal-monitor dashboard: two phase-shifted waveforms plotted over a
// gridded chart, a level meter beside it, and a legend. It is deliberately a
// LAYOUT rather than a scribble — every primitive the adapter maps is exercised
// somewhere a viewer can check it: SDL_FillRect for panels/bars/swatches,
// SDL_DrawLine for the grid, axes and curves, SDL_DrawText for labels, and
// SDL_MapRGB for every colour.
//
// No math.h: the waveform reads a 36-entry sine table (10-degree steps, scaled
// by 1000), which keeps the sample self-contained and integer-only.

#include <stdint.h>

typedef uint32_t Uint32;
typedef int32_t Sint32;
typedef void *SDL_Window;
typedef void *SDL_Surface;

typedef struct SDL_Rect {
    Sint32 x;
    Sint32 y;
    Sint32 w;
    Sint32 h;
} SDL_Rect;

#define SDL_INIT_VIDEO 0x00000020
#define SDL_WINDOW_SHOWN 0x00000004

extern int SDL_Init(Uint32 flags);
extern void SDL_Quit(void);
extern SDL_Window *SDL_CreateWindow(
    const char *title,
    Sint32 x,
    Sint32 y,
    Sint32 w,
    Sint32 h,
    Uint32 flags
);
extern int SDL_DestroyWindow(SDL_Window *window);
extern SDL_Surface *SDL_GetWindowSurface(SDL_Window *window);
extern int SDL_FillRect(SDL_Surface *surface, const SDL_Rect *rect, Uint32 color);
extern int SDL_UpdateWindowSurface(SDL_Window *window);
extern void SDL_Delay(Uint32 ms);
extern Uint32 SDL_MapRGB(Uint32 fmt, Uint32 r, Uint32 g, Uint32 b);
extern int SDL_DrawText(SDL_Surface *surface, const char *text, Sint32 x, Sint32 y, Uint32 color);
extern int SDL_DrawLine(SDL_Surface *surface, Sint32 x1, Sint32 y1, Sint32 x2, Sint32 y2, Uint32 color);
extern int SDL_ShowWindow(SDL_Window *window);

// sin(deg) * 1000, 10-degree steps.
static const Sint32 SINE[36] = {
       0,  174,  342,  500,  643,  766,  866,  940,  985, 1000,
     985,  940,  866,  766,  643,  500,  342,  174,    0, -174,
    -342, -500, -643, -766, -866, -940, -985, -1000, -985, -940,
    -866, -766, -643, -500, -342, -174
};

#define WIN_W 800
#define WIN_H 480

#define CHART_X 40
#define CHART_Y 96
#define CHART_W 480
#define CHART_H 288

#define METER_X 560
#define METER_Y 96
#define METER_W 200
#define METER_H 288

static void fill(SDL_Surface *s, Sint32 x, Sint32 y, Sint32 w, Sint32 h, Uint32 c) {
    SDL_Rect r = {x, y, w, h};
    SDL_FillRect(s, &r, c);
}

// Outline drawn as four lines — the adapter maps DrawLine, not DrawRect.
static void frame_rect(SDL_Surface *s, Sint32 x, Sint32 y, Sint32 w, Sint32 h, Uint32 c) {
    SDL_DrawLine(s, x,     y,     x + w, y,     c);
    SDL_DrawLine(s, x + w, y,     x + w, y + h, c);
    SDL_DrawLine(s, x + w, y + h, x,     y + h, c);
    SDL_DrawLine(s, x,     y + h, x,     y,     c);
}

// One waveform, sampled every 8px and joined with line segments.
static void plot_wave(SDL_Surface *s, Sint32 phase, Sint32 amplitude, Uint32 colour) {
    Sint32 mid = CHART_Y + CHART_H / 2;
    Sint32 prev_x = CHART_X;
    Sint32 prev_y = mid;
    for (Sint32 px = 0; px <= CHART_W; px = px + 8) {
        Sint32 idx = ((px / 8) + phase) % 36;
        Sint32 y = mid - (SINE[idx] * amplitude) / 1000;
        Sint32 x = CHART_X + px;
        if (px > 0) {
            SDL_DrawLine(s, prev_x, prev_y, x, y, colour);
        }
        prev_x = x;
        prev_y = y;
    }
}

int main(void) {
    if (SDL_Init(SDL_INIT_VIDEO) != 0) {
        return 1;
    }

    SDL_Window *window = SDL_CreateWindow("Vybe SDL Adapter - Signal Monitor",
                                          100, 100, WIN_W, WIN_H, SDL_WINDOW_SHOWN);
    if (window == (void *)0) {
        SDL_Quit();
        return 1;
    }

    SDL_Surface *surface = SDL_GetWindowSurface(window);
    if (surface == (void *)0) {
        SDL_DestroyWindow(window);
        SDL_Quit();
        return 1;
    }

    SDL_ShowWindow(window);

    Uint32 bg      = SDL_MapRGB(0,  24,  28,  38);
    Uint32 panel   = SDL_MapRGB(0,  32,  38,  52);
    Uint32 header  = SDL_MapRGB(0,  46,  54,  74);
    Uint32 grid    = SDL_MapRGB(0,  72,  82, 106);
    Uint32 axis    = SDL_MapRGB(0, 110, 122, 148);
    Uint32 ink     = SDL_MapRGB(0, 232, 238, 248);
    Uint32 muted   = SDL_MapRGB(0, 150, 162, 186);
    Uint32 cyan    = SDL_MapRGB(0,  86, 208, 224);
    Uint32 amber   = SDL_MapRGB(0, 240, 176,  80);
    Uint32 green   = SDL_MapRGB(0, 122, 208, 128);

    // NOTE ON TIMING: vybe:gui presents AFTER the program returns —
    // SDL_UpdateWindowSurface marks the frame boundary but does not pump the
    // window, and SDL_Delay blocks the thread the UI needs. So a long loop just
    // stalls with a spinner and only the LAST frame is ever shown. Keep the
    // loop short until per-frame presentation exists.
    for (Sint32 frame = 0; frame < 6; frame = frame + 1) {
        // ── Background and header ──────────────────────────────────────────
        fill(surface, 0, 0, WIN_W, WIN_H, bg);
        fill(surface, 0, 0, WIN_W, 56, header);
        SDL_DrawText(surface, "Signal Monitor", 24, 18, ink);
        SDL_DrawText(surface, "Vybe SDL adapter demo", 560, 22, muted);

        // ── Chart panel ────────────────────────────────────────────────────
        fill(surface, CHART_X, CHART_Y, CHART_W, CHART_H, panel);
        frame_rect(surface, CHART_X, CHART_Y, CHART_W, CHART_H, axis);

        // Horizontal gridlines at quarters, vertical every 60px.
        for (Sint32 g = 1; g < 4; g = g + 1) {
            Sint32 gy = CHART_Y + (CHART_H * g) / 4;
            SDL_DrawLine(surface, CHART_X + 1, gy, CHART_X + CHART_W - 1, gy, grid);
        }
        for (Sint32 gx = CHART_X + 60; gx < CHART_X + CHART_W; gx = gx + 60) {
            SDL_DrawLine(surface, gx, CHART_Y + 1, gx, CHART_Y + CHART_H - 1, grid);
        }
        // Zero axis, brighter than the grid.
        SDL_DrawLine(surface, CHART_X + 1, CHART_Y + CHART_H / 2,
                     CHART_X + CHART_W - 1, CHART_Y + CHART_H / 2, axis);

        // ── Waveforms ──────────────────────────────────────────────────────
        plot_wave(surface, frame, 118, cyan);
        plot_wave(surface, frame + 9, 74, amber);

        SDL_DrawText(surface, "channel A / B", CHART_X + 8, CHART_Y + 8, muted);

        // ── Level meter ────────────────────────────────────────────────────
        fill(surface, METER_X, METER_Y, METER_W, METER_H, panel);
        frame_rect(surface, METER_X, METER_Y, METER_W, METER_H, axis);
        SDL_DrawText(surface, "levels", METER_X + 8, METER_Y + 8, muted);

        for (Sint32 b = 0; b < 5; b = b + 1) {
            Sint32 idx = (frame * 2 + b * 7) % 36;
            Sint32 mag = SINE[idx];
            if (mag < 0) {
                mag = -mag;
            }
            Sint32 bar_h = 30 + (mag * 196) / 1000;
            Sint32 bx = METER_X + 22 + b * 32;
            Sint32 by = METER_Y + METER_H - 24 - bar_h;
            fill(surface, bx, by, 20, bar_h, b % 2 == 0 ? cyan : green);
        }
        SDL_DrawLine(surface, METER_X + 12, METER_Y + METER_H - 24,
                     METER_X + METER_W - 12, METER_Y + METER_H - 24, axis);

        // ── Legend ─────────────────────────────────────────────────────────
        fill(surface, CHART_X, 412, 12, 12, cyan);
        SDL_DrawText(surface, "channel A", CHART_X + 20, 410, ink);
        fill(surface, CHART_X + 130, 412, 12, 12, amber);
        SDL_DrawText(surface, "channel B", CHART_X + 150, 410, ink);
        fill(surface, CHART_X + 260, 412, 12, 12, green);
        SDL_DrawText(surface, "level", CHART_X + 280, 410, ink);

        SDL_DrawText(surface, "FillRect / DrawLine / DrawText / MapRGB", CHART_X, 444, muted);

        SDL_UpdateWindowSurface(window);
    }

    // Deliberately NOT calling SDL_DestroyWindow: it maps to closeForm, which
    // sets `close_requested`, and the real event loop only starts once this
    // function returns — the window would close before painting anything.
    SDL_Quit();
    return 0;
}
