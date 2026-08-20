/* Headless self-test: injects the input stream through SDL_PushEvent —
 * the same queue and handler the interactive loop uses — and asserts
 * the engine's answers. Run with: make test  (vybex --entry self_test) */
#include <SDL2/SDL.h>
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include "calc.h"
#include "ui.h"

static void push_key(int sym) {
    SDL_Event e;
    e.type = SDL_KEYDOWN;
    e.key.keysym.sym = sym;
    SDL_PushEvent(&e);
}

static void push_click(int x, int y) {
    SDL_Event e;
    e.type = SDL_MOUSEBUTTONDOWN;
    e.button.button = SDL_BUTTON_LEFT;
    e.button.x = x;
    e.button.y = y;
    SDL_PushEvent(&e);
}

static void drain(struct Calc *c) {
    SDL_Event e;
    while (SDL_PollEvent(&e)) {
        ui_handle_event(c, &e);
    }
}

static void expect(const struct Calc *c, const char *want, const char *what) {
    char got[64];
    calc_display(c, got, sizeof(got));
    if (strcmp(got, want) != 0) {
        printf("FAIL %s: got [%s] want [%s]\n", what, got, want);
        assert(0);
    }
    printf("ok %s = %s\n", what, got);
}

int self_test(void) {
    struct Calc c;
    calc_reset(&c);

    /* keyboard: 12 + 34 = 46 */
    push_key(SDLK_1); push_key(SDLK_2);
    push_key(SDLK_PLUS);
    push_key(SDLK_3); push_key(SDLK_4);
    push_key(SDLK_RETURN);
    drain(&c);
    expect(&c, "46", "12+34");

    /* keyboard: sqrt(81) via 'c' clear, 8 1, s */
    push_key(SDLK_c);
    push_key(SDLK_8); push_key(SDLK_1);
    push_key(SDLK_s);
    drain(&c);
    expect(&c, "9", "sqrt 81");

    /* mouse: C, 7 * 6 = 42 — click the button centers.
     * cell_w=50 cell_h=41 top=72: center(col,row) = (33+col*58, 92+row*49).
     * layout: row0 "C..r" row1 "789/" row2 "456*" row3 "123-" row4 "0=.+" */
    push_click(33, 92);    /* C */
    push_click(33, 141);   /* 7 */
    push_click(207, 190);  /* * */
    push_click(149, 190);  /* 6 */
    push_click(91, 288);   /* = */
    drain(&c);
    expect(&c, "42", "7*6 by mouse");

    /* every remaining control, keyboard: 9-5=4, 7/2=3.5, decimals */
    push_key(SDLK_c);
    push_key(SDLK_9); push_key(SDLK_MINUS); push_key(SDLK_5); push_key(SDLK_RETURN);
    drain(&c);
    expect(&c, "4", "9-5");

    push_key(SDLK_c);
    push_key(SDLK_7); push_key(SDLK_SLASH); push_key(SDLK_2); push_key(SDLK_EQUALS);
    drain(&c);
    expect(&c, "3.5", "7/2");

    push_key(SDLK_c);
    push_key(SDLK_1); push_key(SDLK_PERIOD); push_key(SDLK_5);
    push_key(SDLK_PLUS);
    push_key(SDLK_2); push_key(SDLK_PERIOD); push_key(SDLK_2); push_key(SDLK_5);
    push_key(SDLK_RETURN);
    drain(&c);
    expect(&c, "3.75", "1.5+2.25");

    /* decimal + division by mouse: 1 . 5 * 2 = 3  (row4: "0=.+" → '.' col2) */
    push_click(33, 92);    /* C */
    push_click(33, 237);   /* 1 (row 3 col 0) */
    push_click(149, 288);  /* . (row 4 col 2) */
    push_click(149, 190);  /* 5? no — 5 is row 2 col 1 */
    drain(&c);
    calc_reset(&c);
    push_click(33, 237);   /* 1 */
    push_click(149, 288);  /* . */
    push_click(91, 190);   /* 5 (row 2 col 1) */
    push_click(207, 190);  /* * */
    push_click(91, 141);   /* 8 (row 1 col 1) */
    push_click(91, 288);   /* = */
    drain(&c);
    expect(&c, "12", "1.5*8 by mouse");

    printf("selftest passed\n");
    return 0;
}
