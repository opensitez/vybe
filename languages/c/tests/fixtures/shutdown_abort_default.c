#include <stdio.h>
#include <stdlib.h>
static void cleanup(void) { printf("BAD\n"); }
int main(void) { atexit(cleanup); printf("unflushed"); abort(); }
