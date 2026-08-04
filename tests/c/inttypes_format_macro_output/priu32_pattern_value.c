// vybe-test: c/inttypes_format_macro_output/priu32_pattern_value
// origin: languages/c/tests/c/test_inttypes_format_macro_output.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <inttypes.h>
int main() {
const char *__w[] = {"305419896\n"};
int __n = 1, __i = 0;
uint32_t v=305419896U; { char __t[512]; snprintf(__t, sizeof(__t), "%" PRIu32 "\n", v);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

