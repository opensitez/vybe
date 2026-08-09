// vybe-test: c/cover_uchar_h/uchar_utf16_pointer_reads_units
// origin: languages/c/tests/c/test_cover_uchar_h.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <uchar.h>
int main() {
const char *__w[] = {"104 233 55348 56606\n"};
int __n = 1, __i = 0;
const char16_t *s = u"hé\U0001D11E"; { char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d %d\n", (int)s[0], (int)s[1], (int)s[2], (int)s[3]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}
