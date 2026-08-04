// vybe-test: c/lang_array_decay_parameters/incomplete_array_param_declaration_compiles
// origin: languages/c/tests/c/test_lang_array_decay_parameters.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
void touch(int a[]); void touch(int a[]){ a[0]=1; }
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
int a[1]={0}; touch(a); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", a[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

