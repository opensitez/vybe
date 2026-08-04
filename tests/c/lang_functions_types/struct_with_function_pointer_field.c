// vybe-test: c/lang_functions_types/struct_with_function_pointer_field
// origin: languages/c/tests/c/test_lang_functions_types.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
typedef struct { int (*op)(int); } VTable; int neg(int x){return -x;}
int main() {
const char *__w[] = {"-5\n"};
int __n = 1, __i = 0;
VTable vt={neg}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", vt.op(5));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

