// vybe-test: c/c_atomic_basic_types/atomic_int_init_load
// origin: languages/c/tests/c/test_c_atomic_basic_types.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdatomic.h>
int main() {const char *__w[] = {"42"};
int __n = 1, __i = 0;
 atomic_int val = ATOMIC_VAR_INIT(42); { char __t[512]; snprintf(__t, sizeof(__t), "%d", atomic_load(&val));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

