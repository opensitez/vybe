#include <stdio.h>
#include <stdlib.h>
static void late(void) { printf("L"); }
static void first(void) { printf("A"); }
static void last(void) { printf("B"); atexit(late); }
static void skipped(void) { printf("BAD"); }
static void status_handler(int status, void *arg) {
    printf("%d:%d", status, *(int *)arg);
}
static int payload = 42;
static int other_payload = 99;
static void register_local(void) {
    void (*handler)(void) = first;
    int *argument = &payload;
    atexit(handler);
    for (int i = 0; i < 2; ++i) atexit(first);
    handler = skipped;
    if (0) atexit(skipped);
    on_exit(status_handler, argument);
    argument = &other_payload;
    if (atexit(last) != 0) abort();
}
int main(void) {
    register_local();
    printf("P");
    return 7;
}
