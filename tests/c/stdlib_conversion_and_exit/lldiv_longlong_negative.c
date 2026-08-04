// vybe-test: c/stdlib_conversion_and_exit/lldiv_longlong_negative
// origin: languages/c/tests/c/test_stdlib_conversion_and_exit.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <stdlib.h>
int main() {
const char *__w[] = {"-6 -1\n"};
int __n = 1, __i = 0;
lldiv_t r = lldiv(-25LL, 4LL); { char __t[512]; snprintf(__t, sizeof(__t), "%lld %lld\n", r.quot, r.rem);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

