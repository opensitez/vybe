// vybe-test: c/string_bounded_copy_concat/strncat_uppercase_chunk
// origin: languages/c/tests/c/test_string_bounded_copy_concat.rs
#include <assert.h>
#include <stdio.h>
#include <string.h>
char d[12] = "mix";
int main() {
const char *__w[] = {"mixED\n"};
int __n = 1, __i = 0;
strncat(d, "ED", 2); { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", d);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

