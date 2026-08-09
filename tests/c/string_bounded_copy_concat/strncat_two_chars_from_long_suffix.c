// vybe-test: c/string_bounded_copy_concat/strncat_two_chars_from_long_suffix
// origin: languages/c/tests/c/test_string_bounded_copy_concat.rs
#include <assert.h>
#include <stdio.h>
#include <string.h>
char d[12] = "q";
int main() {
const char *__w[] = {"qwe\n"};
int __n = 1, __i = 0;
strncat(d, "werty", 2); { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", d);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

