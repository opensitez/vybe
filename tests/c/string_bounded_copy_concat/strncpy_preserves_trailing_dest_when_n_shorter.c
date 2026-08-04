// vybe-test: c/string_bounded_copy_concat/strncpy_preserves_trailing_dest_when_n_shorter
// origin: languages/c/tests/c/test_string_bounded_copy_concat.rs
#include <assert.h>
#include <stdio.h>
#include <string.h>
char d[6] = "abcdef";
int main() {
const char *__w[] = {"f\n"};
int __n = 1, __i = 0;
strncpy(d, "12", 2); { char __t[512]; snprintf(__t, sizeof(__t), "%c\n", d[5]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

