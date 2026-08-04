// vybe-test: c/lang_logical_short_circuit/logical_result_stored_in_int
// origin: languages/c/tests/c/test_lang_logical_short_circuit.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"1 1\n"};
int __n = 1, __i = 0;
int a=2&&3, b=0||4; { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", a, b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

