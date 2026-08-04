// vybe-test: c/memory_functions/realloc_grows_buffer
// origin: languages/c/tests/c/test_memory_functions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"1 2 3 4\n"};
int __n = 1, __i = 0;

int *p = (int*)malloc(2 * sizeof(int));
p[0] = 1; p[1] = 2;
p = (int*)realloc(p, 4 * sizeof(int));
p[2] = 3; p[3] = 4;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d %d\n", p[0], p[1], p[2], p[3]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
free(p);
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

