#include <signal.h>
#include <stdio.h>
#include <stdlib.h>
static void cleanup(void) { printf("H"); }
static void exiting_handler(int sig) {
    if (sig != SIGABRT) _Exit(99);
    printf("P");
    exit(8);
}
int main(void) {
    atexit(cleanup);
    signal(SIGABRT, exiting_handler);
    abort();
    printf("BAD\n");
}
