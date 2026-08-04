// vybe-test: c/static_assert/static_assert_in_function_scope
// origin: languages/c/tests/c/test_static_assert.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int main() {const char *__w[] = {"ok\n"};
int __n = 1, __i = 0;

    _Static_assert(sizeof(long) >= sizeof(int), "long must be at least int size");
    { char __t[512]; snprintf(__t, sizeof(__t), "ok\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

