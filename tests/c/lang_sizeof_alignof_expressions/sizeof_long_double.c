// vybe-test: c/lang_sizeof_alignof_expressions/sizeof_long_double
// origin: languages/c/tests/c/test_lang_sizeof_alignof_expressions.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
char exp_buf[16]; snprintf(exp_buf, sizeof(exp_buf), "%d\n", (int)sizeof(long double));
const char *__w[] = {exp_buf};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (int)sizeof(long double));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

