#include <stdio.h>
#include <stdlib.h>
static void handler(void) { printf("BAD\n"); }
int main(void) { atexit(handler); printf("partial"); _Exit(4); abort(); }
