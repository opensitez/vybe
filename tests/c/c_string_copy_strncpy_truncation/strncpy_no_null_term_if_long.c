// vybe-test: c/c_string_copy_strncpy_truncation/strncpy_no_null_term_if_long
// origin: languages/c/tests/c/test_c_string_copy_strncpy_truncation.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <string.h>
int main() {const char *__w[] = {"hello"};
int __n = 1, __i = 0;
 char dest[5] = "xxxxx"; strncpy(dest, "hello", 5); { char __t[512]; snprintf(__t, sizeof(__t), "%c%c%c%c%c", dest[0], dest[1], dest[2], dest[3], dest[4]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

