// vybe-test: c/string_bounded_copy_concat/strncpy_truncates_long_source_at_n
// origin: languages/c/tests/c/test_string_bounded_copy_concat.rs
#include <assert.h>
#include <stdio.h>
#include <string.h>
int main() {
const char *__w[] = {"gala\n"};
int __n = 1, __i = 0;
char d[5]; strncpy(d, "galaxy", 4); { char __t[512]; snprintf(__t, sizeof(__t), "%c%c%c%c\n", d[0],d[1],d[2],d[3]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

