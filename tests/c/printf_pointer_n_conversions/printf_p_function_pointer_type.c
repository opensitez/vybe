// vybe-test: c/printf_pointer_n_conversions/printf_p_function_pointer_type
// origin: languages/c/tests/c/test_printf_pointer_n_conversions.rs
#include <assert.h>
#include <stdio.h>
#include <stddef.h>
#include <string.h>
int id(int v){ return v; }
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
char a[64],b[64]; sprintf(a,"%p",(void*)id); sprintf(b,"%p",(void*)id); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", strcmp(a,b)==0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

