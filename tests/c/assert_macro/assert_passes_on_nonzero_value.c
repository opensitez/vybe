// vybe-test: c/assert_macro/assert_passes_on_nonzero_value
// origin: languages/c/tests/c/test_assert_macro.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <assert.h>
int main() {const char *__w[] = {"passed\n"};
int __n = 1, __i = 0;

    assert(42);
    { char __t[512]; snprintf(__t, sizeof(__t), "passed\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

