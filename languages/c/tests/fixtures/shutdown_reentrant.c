#include <stdio.h>
#include <stdlib.h>
static void remaining(void) { printf("BAD"); }
static void handler(void) { printf("H"); exit(9); printf("BAD"); }
int main(void) { atexit(remaining); atexit(handler); printf("P"); exit(7); abort(); }
