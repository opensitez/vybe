#include <math.h>
#include <stdio.h>
#include "calc.h"

void calc_reset(struct Calc *c) {
    c->accum = 0;
    c->entry = 0;
    c->entering = 0;
    c->frac = 0;
    c->pending = 0;
}

void calc_digit(struct Calc *c, int d) {
    if (!c->entering) { c->entry = 0; c->entering = 1; c->frac = 0; }
    if (c->frac > 0) {
        c->entry = c->entry + d * c->frac;
        c->frac = c->frac / 10.0;
    } else {
        c->entry = c->entry * 10 + d;
    }
}

void calc_dot(struct Calc *c) {
    if (!c->entering) { c->entry = 0; c->entering = 1; }
    if (c->frac == 0) { c->frac = 0.1; }
}

static void apply_pending(struct Calc *c) {
    double rhs = c->entering ? c->entry : c->accum;
    switch (c->pending) {
        case '+': c->accum = c->accum + rhs; break;
        case '-': c->accum = c->accum - rhs; break;
        case '*': c->accum = c->accum * rhs; break;
        case '/': c->accum = rhs != 0 ? c->accum / rhs : 0; break;
        default:  c->accum = rhs; break;
    }
}

void calc_op(struct Calc *c, char op) {
    apply_pending(c);
    c->pending = op;
    c->entering = 0;
    c->frac = 0;
}

void calc_equals(struct Calc *c) {
    apply_pending(c);
    c->pending = 0;
    c->entering = 0;
    c->frac = 0;
}

void calc_sqrt(struct Calc *c) {
    calc_equals(c);
    c->accum = sqrt(c->accum);
}

double calc_value(const struct Calc *c) {
    return c->entering ? c->entry : c->accum;
}
