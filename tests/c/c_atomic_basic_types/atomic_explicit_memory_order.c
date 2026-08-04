// vybe-test: c/c_atomic_basic_types/atomic_explicit_memory_order
// origin: languages/c/tests/c/test_c_atomic_basic_types.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdatomic.h>
int main() {const char *__w[] = {"5"};
int __n = 1, __i = 0;
 atomic_int val = 0; atomic_store_explicit(&val, 5, memory_order_release); { char __t[512]; snprintf(__t, sizeof(__t), "%d", atomic_load_explicit(&val, memory_order_acquire));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

