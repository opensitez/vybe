// vybe-test: c/string_copy_concat/strstr_can_find_middle_segment_after_copy
// origin: languages/c/tests/c/test_string_copy_concat.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
char dest[32] = "";
int main() {
const char *__w[] = {"narama\n"};
int __n = 1, __i = 0;
strcpy(dest, "bananarama"); { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", strstr(dest, "nara"));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

