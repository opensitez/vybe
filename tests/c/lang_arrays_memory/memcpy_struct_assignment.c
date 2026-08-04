// vybe-test: c/lang_arrays_memory/memcpy_struct_assignment
// origin: languages/c/tests/c/test_lang_arrays_memory.rs
#include <assert.h>
#include <stdio.h>
#include <string.h>
struct P { int x; };
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
struct P a={1},b; b=a; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", b.x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

