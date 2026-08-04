// vybe-test: c/functions_advanced/function_can_return_struct_like_multiple_outputs_via_pointers
// origin: languages/c/tests/c/test_functions_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
void divmod(int a, int b, int *q, int *r) { *q = a / b; *r = a % b; }
int main() {
const char *__w[] = {"3 2\n"};
int __n = 1, __i = 0;
int q = 0; int r = 0;
divmod(17, 5, &q, &r);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", q, r);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

