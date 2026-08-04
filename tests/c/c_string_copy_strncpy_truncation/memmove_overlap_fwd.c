// vybe-test: c/c_string_copy_strncpy_truncation/memmove_overlap_fwd
// origin: languages/c/tests/c/test_c_string_copy_strncpy_truncation.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <string.h>
int main() {const char *__w[] = {"ababcd"};
int __n = 1, __i = 0;
 char s[] = "abcdef"; memmove(s + 2, s, 4); { char __t[512]; snprintf(__t, sizeof(__t), "%s", s);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

