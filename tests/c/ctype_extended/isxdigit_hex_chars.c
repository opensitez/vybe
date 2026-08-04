// vybe-test: c/ctype_extended/isxdigit_hex_chars
// origin: languages/c/tests/c/test_ctype_extended.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"1 1 1 0\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d %d\n", isxdigit('0') != 0, isxdigit('a') != 0, isxdigit('F') != 0, isxdigit('g') != 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

