// vybe-test: c/string_copy_concat/copy_then_concat_then_compare_can_equal_expected
// origin: languages/c/tests/c/test_string_copy_concat.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
char dest[32] = "";
int main() {
const char *__w[] = {"0\n"};
int __n = 1, __i = 0;
strcpy(dest, "hello"); strcat(dest, "world"); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", strcmp(dest, "helloworld"));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

