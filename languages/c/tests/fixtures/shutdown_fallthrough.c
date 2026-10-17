#include <stdio.h>
#include <stdlib.h>
static void handler(void) { printf("H"); }
int main(void) { atexit(handler); printf("P"); }
