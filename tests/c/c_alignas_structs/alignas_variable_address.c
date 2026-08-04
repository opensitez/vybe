// vybe-test: c/c_alignas_structs/alignas_variable_address
// origin: languages/c/tests/c/test_c_alignas_structs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdalign.h>
#include <stdint.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 alignas(32) int x; { char __t[512]; snprintf(__t, sizeof(__t), "%d", ((uintptr_t)&x % 32) == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

