// vybe-test: c/cover_uchar_h/uchar_sizeof_types_wasm32
// origin: languages/c/tests/c/test_cover_uchar_h.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <uchar.h>
int main() {
const char *__w[] = {"1 2 4\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d\n", (int)sizeof(char8_t), (int)sizeof(char16_t), (int)sizeof(char32_t));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

