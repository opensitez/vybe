// vybe-test: c/lang_logical_short_circuit/or_empty_rhs_not_evaluated
// origin: languages/c/tests/c/test_lang_logical_short_circuit.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int x=1;
int main() {
const char *__w[] = {"1 1\n"};
int __n = 1, __i = 0;
int r = 1 || (x=9); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", r, x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

