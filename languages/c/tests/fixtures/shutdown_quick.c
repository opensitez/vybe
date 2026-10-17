#include <stdio.h>
#include <stdlib.h>
static void regular(void) { printf("BAD\n"); }
static void late(void) { printf("L\n"); printf("unflushed"); }
static void quick(void) { printf("Q\n"); at_quick_exit(late); }
int main(void) {
    atexit(regular);
    at_quick_exit(quick);
    quick_exit(3);
    abort();
}
