#include <signal.h>
#include <stdio.h>
#include <stdlib.h>
static void cleanup(void) { printf("BAD\n"); }
static void returning_handler(int sig) {
    if (sig != SIGABRT) _Exit(99);
    printf("signal\n");
    printf("unflushed");
}
int main(void) {
    atexit(cleanup);
    signal(SIGABRT, returning_handler);
    abort();
    printf("BAD\n");
}
