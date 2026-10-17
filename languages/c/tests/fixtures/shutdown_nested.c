#include <stdio.h>
#include <stdlib.h>
static void handler(void) { printf("H"); }
static void nested(void) { printf("P"); exit(7); printf("BAD"); }
int main(void) { atexit(handler); nested(); abort(); }
