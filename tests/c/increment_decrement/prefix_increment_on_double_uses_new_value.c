// vybe-test: c/increment_decrement/prefix_increment_on_double_uses_new_value
// origin: languages/c/tests/c/test_increment_decrement.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
double x = 1.5;
int main() {
const char *__w[] = {"2.5\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%.1f\n", ++x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

