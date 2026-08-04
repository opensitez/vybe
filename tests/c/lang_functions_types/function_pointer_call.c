// vybe-test: c/lang_functions_types/function_pointer_call
// origin: languages/c/tests/c/test_lang_functions_types.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int inc(int x){return x+1;}
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;
int (*fp)(int)=inc; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", fp(2));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

