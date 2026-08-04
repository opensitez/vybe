// vybe-test: c/nested_functions/function_modifying_via_pointer
// origin: languages/c/tests/c/test_nested_functions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
void double_val(int *p) { *p *= 2; }
int main() {
const char *__w[] = {"14\n"};
int __n = 1, __i = 0;
int x = 7;
double_val(&x);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

