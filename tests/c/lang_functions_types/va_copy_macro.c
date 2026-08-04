// vybe-test: c/lang_functions_types/va_copy_macro
// origin: languages/c/tests/c/test_lang_functions_types.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <stdarg.h>
int first(int n,...){ va_list ap,ap2; va_start(ap,n); va_copy(ap2,ap); int a=va_arg(ap,int); int b=va_arg(ap2,int); va_end(ap); va_end(ap2); return a+b; }
int main() {
const char *__w[] = {"8\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", first(2,4,5));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

