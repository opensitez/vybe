#ifndef UI_H
#define UI_H
#include <SDL2/SDL.h>
#include "calc.h"

#define CALC_W 240
#define CALC_H 320

/* Map a mouse click to a key label ('0'-'9','+','-','*','/','=','r','C'),
 * or 0 for no button. 'r' is the sqrt key. */
char ui_hit_test(int x, int y);

/* Feed one SDL event into the calculator. Returns 0 when the event asks to
 * quit, 1 otherwise. Shared by the interactive loop and the headless test. */
int ui_handle_event(struct Calc *c, SDL_Event *e);

/* Apply one key label to the engine. */
void ui_press(struct Calc *c, char key);

void ui_render(SDL_Surface *screen, const struct Calc *c);

#endif
