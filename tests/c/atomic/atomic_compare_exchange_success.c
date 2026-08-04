// vybe-test: c/atomic/atomic_compare_exchange_success
// origin: languages/c/tests/c/test_atomic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <stdatomic.h>
int main() {const char *__w[] = {"1 10\n"};
int __n = 1, __i = 0;

    atomic_int x = 5;
    int expected = 5;
    int ok = atomic_compare_exchange_strong(&x, &expected, 10);
    { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", ok, atomic_load(&x));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

