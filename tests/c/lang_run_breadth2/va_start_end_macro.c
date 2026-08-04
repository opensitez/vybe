// vybe-test: c/lang_run_breadth2/va_start_end_macro
// origin: languages/c/tests/c/test_lang_run_breadth2.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <stdarg.h>
int add2(int n,...){va_list ap; va_start(ap,n); int a=va_arg(ap,int); int b=va_arg(ap,int); va_end(ap); return a+b;}
int main() {
const char *__w[] = {"7\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", add2(2,3,4));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

