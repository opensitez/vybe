#include <signal.h>
#include <stdio.h>
#include <stdlib.h>
static void cleanup(void) { printf("BAD\n"); }
int main(void) {
    atexit(cleanup);
    signal(SIGABRT, SIG_IGN);
    printf("unflushed");
    (abort(), printf("BAD\n"));
}
