// vybe-test: c/string_copy_concat/strcpy_overwrites_old_suffix_completely
// origin: languages/c/tests/c/test_string_copy_concat.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
char dest[32] = "abcdef";
int main() {
const char *__w[] = {"xy\n"};
int __n = 1, __i = 0;
strcpy(dest, "xy"); { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", dest);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

