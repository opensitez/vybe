// vybe-test: c/string_stdlib/strncat_appends_n_chars
// origin: languages/c/tests/c/test_string_stdlib.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"hello world\n"};
int __n = 1, __i = 0;

char dst[20] = "hello";
strncat(dst, " world extra", 6);
{ char __t[512]; snprintf(__t, sizeof(__t), "%s\n", dst);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

