// vybe-test: c/lang_cast_expression_semantics/double_pointer_to_void_and_back
// origin: languages/c/tests/c/test_lang_cast_expression_semantics.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"3.5\n"};
int __n = 1, __i = 0;
double d = 3.5; void *vp = (void *)&d; { char __t[512]; snprintf(__t, sizeof(__t), "%.1f\n", *(double *)vp);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

