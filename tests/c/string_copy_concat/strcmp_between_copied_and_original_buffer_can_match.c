// vybe-test: c/string_copy_concat/strcmp_between_copied_and_original_buffer_can_match
// origin: languages/c/tests/c/test_string_copy_concat.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
char left[32] = "copy"; char right[32] = "";
int main() {
const char *__w[] = {"0\n"};
int __n = 1, __i = 0;
strcpy(right, left); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", strcmp(left, right));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

