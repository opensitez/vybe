#include <SDL2/SDL.h>
#include <stdio.h>
#include "ui.h"

/* 4x5 button grid under a display strip. */
static const char *LAYOUT[5] = { "C  r", "789/", "456*", "123-", "0=.+" };

#define PAD 8
#define DISPLAY_H 56
#define COLS 4
#define ROWS 5

static int cell_w(void) { return (CALC_W - PAD * (COLS + 1)) / COLS; }
static int cell_h(void) { return (CALC_H - DISPLAY_H - PAD * (ROWS + 2)) / ROWS; }

char ui_hit_test(int x, int y) {
    int cw = cell_w(), ch = cell_h();
    int top = DISPLAY_H + PAD * 2;
    int col, row;
    if (y < top) return 0;
    col = (x - PAD) / (cw + PAD);
    row = (y - top) / (ch + PAD);
    if (col < 0 || col >= COLS || row < 0 || row >= ROWS) return 0;
    {
        char k = LAYOUT[row][col];
        if (k == ' ') return 0;
        return k;
    }
}

void ui_press(struct Calc *c, char key) {
    if (key >= '0' && key <= '9') { calc_digit(c, key - '0'); return; }
    switch (key) {
        case '.': calc_dot(c); break;
        case '+': case '-': case '*': case '/': calc_op(c, key); break;
        case '=': calc_equals(c); break;
        case 'r': calc_sqrt(c); break;
        case 'C': calc_reset(c); break;
        default: break;
    }
}

int ui_handle_event(struct Calc *c, SDL_Event *e) {
    if (e->type == SDL_QUIT) return 0;
    if (e->type == SDL_MOUSEBUTTONDOWN) {
        char k = ui_hit_test(e->button.x, e->button.y);
        if (k) ui_press(c, k);
        return 1;
    }
    if (e->type == SDL_KEYDOWN) {
        int sym = e->key.keysym.sym;
        if (sym >= SDLK_0 && sym <= SDLK_9) { calc_digit(c, sym - SDLK_0); return 1; }
        if (sym == SDLK_PERIOD) { calc_dot(c); return 1; }
        if (sym == SDLK_PLUS) { calc_op(c, '+'); return 1; }
        if (sym == SDLK_MINUS) { calc_op(c, '-'); return 1; }
        if (sym == SDLK_ASTERISK) { calc_op(c, '*'); return 1; }
        if (sym == SDLK_SLASH) { calc_op(c, '/'); return 1; }
        if (sym == SDLK_RETURN || sym == SDLK_EQUALS) { calc_equals(c); return 1; }
        if (sym == SDLK_s) { calc_sqrt(c); return 1; }
        if (sym == SDLK_c) { calc_reset(c); return 1; }
        if (sym == SDLK_ESCAPE) return 0;
    }
    return 1;
}

void ui_render(SDL_Surface *screen, const struct Calc *c) {
    char text[64];
    SDL_Rect r;
    int cw = cell_w(), ch = cell_h();
    int top = DISPLAY_H + PAD * 2;
    int row, col;

    r.x = 0; r.y = 0; r.w = CALC_W; r.h = CALC_H;
    SDL_FillRect(screen, &r, SDL_MapRGB(0, 28, 30, 34));

    /* display strip */
    r.x = PAD; r.y = PAD; r.w = CALC_W - PAD * 2; r.h = DISPLAY_H;
    SDL_FillRect(screen, &r, SDL_MapRGB(0, 12, 40, 22));
    calc_display(c, text, sizeof(text));
    SDL_DrawText(screen, text, PAD * 2, PAD + DISPLAY_H / 2 - 8, SDL_MapRGB(0, 120, 255, 150));

    for (row = 0; row < ROWS; row++) {
        for (col = 0; col < COLS; col++) {
            char k = LAYOUT[row][col];
            char label[2];
            if (k == ' ') continue;
            r.x = PAD + col * (cw + PAD);
            r.y = top + row * (ch + PAD);
            r.w = cw; r.h = ch;
            SDL_FillRect(screen, &r, k >= '0' && k <= '9'
                ? SDL_MapRGB(0, 58, 62, 70)
                : SDL_MapRGB(0, 90, 70, 40));
            label[0] = k == 'r' ? 'V' : k; /* V for sqrt */
            label[1] = 0;
            SDL_DrawText(screen, label, r.x + cw / 2 - 4, r.y + ch / 2 - 8,
                         SDL_MapRGB(0, 235, 235, 235));
        }
    }
}
