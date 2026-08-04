// vybe-test: c/cover_string_h/strndup_truncates
// origin: languages/c/tests/c/test_cover_string_h.rs
#include <assert.h>
#include <stdio.h>
#include <string.h>
#include <stdlib.h>
int main() {
const char *__w[] = {"abc\n"};
int __n = 1, __i = 0;
char *s = strndup("abcdef", 3); { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", s);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } free(s); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

