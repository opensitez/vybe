// vybe-test: c/lang_logical_short_circuit/and_with_assign_side_effect_once
// origin: languages/c/tests/c/test_lang_logical_short_circuit.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int n=0;
int main() {
const char *__w[] = {"1 2\n"};
int __n = 1, __i = 0;
int r = (n=1) && (n=2); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", r, n);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

