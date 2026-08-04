// vybe-test: c/lang_arrays_memory/union_size_reuse
// origin: languages/c/tests/c/test_lang_arrays_memory.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
union U { int i; double d; };
int main() {
const char *__w[] = {"42\n"};
int __n = 1, __i = 0;
union U u; u.i=42; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", u.i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

