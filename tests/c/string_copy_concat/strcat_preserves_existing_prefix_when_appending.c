// vybe-test: c/string_copy_concat/strcat_preserves_existing_prefix_when_appending
// origin: languages/c/tests/c/test_string_copy_concat.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
char dest[32] = "pre";
int main() {
const char *__w[] = {"prefix\n"};
int __n = 1, __i = 0;
strcat(dest, "fix"); { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", dest);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

