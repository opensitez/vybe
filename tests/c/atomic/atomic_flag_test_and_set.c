// vybe-test: c/atomic/atomic_flag_test_and_set
// origin: languages/c/tests/c/test_atomic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <stdatomic.h>
int main() {const char *__w[] = {"0 1\n"};
int __n = 1, __i = 0;

    atomic_flag f = ATOMIC_FLAG_INIT;
    int first = atomic_flag_test_and_set(&f);
    int second = atomic_flag_test_and_set(&f);
    { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", first, second);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

