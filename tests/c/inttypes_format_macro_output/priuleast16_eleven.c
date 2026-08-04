// vybe-test: c/inttypes_format_macro_output/priuleast16_eleven
// origin: languages/c/tests/c/test_inttypes_format_macro_output.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <inttypes.h>
int main() {
const char *__w[] = {"11\n"};
int __n = 1, __i = 0;
uint_least16_t v=11; { char __t[512]; snprintf(__t, sizeof(__t), "%" PRIuLEAST16 "\n", v);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

