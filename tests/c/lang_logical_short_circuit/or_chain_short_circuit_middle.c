// vybe-test: c/lang_logical_short_circuit/or_chain_short_circuit_middle
// origin: languages/c/tests/c/test_lang_logical_short_circuit.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int c=0;
int main() {
const char *__w[] = {"1 0\n"};
int __n = 1, __i = 0;
int r = 0 || 1 || (c=1); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", r, c);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

