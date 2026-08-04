// vybe-test: c/c_atomic_flag_spinlock/atomic_flag_clear
// origin: languages/c/tests/c/test_c_atomic_flag_spinlock.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdatomic.h>
int main() {const char *__w[] = {"0"};
int __n = 1, __i = 0;
 atomic_flag lock = ATOMIC_FLAG_INIT; atomic_flag_test_and_set(&lock); atomic_flag_clear(&lock); _Bool was_set = atomic_flag_test_and_set(&lock); { char __t[512]; snprintf(__t, sizeof(__t), "%d", was_set);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

