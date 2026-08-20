#ifndef CALC_H
#define CALC_H
#include <stdio.h>

/* Calculator engine — no SDL here; pure state machine over key events. */

struct Calc {
    double accum;      /* left operand / result           */
    double entry;      /* digits being typed              */
    int    entering;   /* 1 while typing a number         */
    double frac;       /* 0 = integer mode; else next digit's weight */
    char   pending;    /* '+', '-', '*', '/' or 0         */
};

void calc_reset(struct Calc *c);
void calc_digit(struct Calc *c, int d);
void calc_dot(struct Calc *c);
void calc_op(struct Calc *c, char op);
void calc_equals(struct Calc *c);
void calc_sqrt(struct Calc *c);                 /* uses math.h            */
double calc_value(const struct Calc *c);

/* Format the current value into a CALLER-LOCAL buffer. A macro on purpose:
 * the snprintf runs directly on the caller's array — no buffer crosses a
 * call boundary (chars-through-params write-back is a known engine gap). */
#define calc_display(c, buf, cap) do { \
    double __v = calc_value(c); \
    if (__v == (long)__v) { snprintf((buf), (cap), "%ld", (long)__v); } \
    else { snprintf((buf), (cap), "%g", __v); } \
} while (0)

#endif
