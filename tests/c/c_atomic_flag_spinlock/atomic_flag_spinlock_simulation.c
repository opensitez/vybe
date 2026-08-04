// vybe-test: c/c_atomic_flag_spinlock/atomic_flag_spinlock_simulation
// origin: languages/c/tests/c/test_c_atomic_flag_spinlock.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdatomic.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 atomic_flag lock = ATOMIC_FLAG_INIT; while (atomic_flag_test_and_set_explicit(&lock, memory_order_acquire)) {} /* locked */ atomic_flag_clear_explicit(&lock, memory_order_release); { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

