#include "counter.h"
/* file-scope static: PRIVATE to this translation unit */
static int calls = 0;
void count_call(void) { calls++; }
int calls_made(void) { return calls; }
