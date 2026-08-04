// vybe-test: c/cover_stdlib_h/ldiv_t_quotient
// origin: languages/c/tests/c/test_cover_stdlib_h.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <stdlib.h>
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;
ldiv_t r = ldiv(10L, 3L); { char __t[512]; snprintf(__t, sizeof(__t), "%ld\n", r.quot);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

