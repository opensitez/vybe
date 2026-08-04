// vybe-test: c/cover_uchar_h/uchar_c32rtomb_ascii
// origin: languages/c/tests/c/test_cover_uchar_h.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <uchar.h>
int main() {
const char *__w[] = {"1 T\n"};
int __n = 1, __i = 0;
char b[4]; size_t n = c32rtomb(b, U'T', 0); { char __t[512]; snprintf(__t, sizeof(__t), "%d %c\n", (int)n, b[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

