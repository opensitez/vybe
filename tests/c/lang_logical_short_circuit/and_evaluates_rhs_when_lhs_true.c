// vybe-test: c/lang_logical_short_circuit/and_evaluates_rhs_when_lhs_true
// origin: languages/c/tests/c/test_lang_logical_short_circuit.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int side=0;
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
if(1 && (side=1)){} { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", side);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

