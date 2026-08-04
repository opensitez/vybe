// vybe-test: c/lang_preprocessor_macro_expansion/macro_expands_in_initializer
// origin: languages/c/tests/c/test_lang_preprocessor_macro_expansion.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#define SEED 4
int main() {
const char *__w[] = {"5\n"};
int __n = 1, __i = 0;
int v=SEED+1; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", v);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

