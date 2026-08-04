// vybe-test: c/string_bounded_copy_concat/strncpy_from_empty_source_zeroes_first_n
// origin: languages/c/tests/c/test_string_bounded_copy_concat.rs
#include <assert.h>
#include <stdio.h>
#include <string.h>
int main() {
const char *__w[] = {"1 1\n"};
int __n = 1, __i = 0;
char d[4]; memset(d, '9', 4); strncpy(d, "", 3); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", d[0]==0, d[1]==0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

