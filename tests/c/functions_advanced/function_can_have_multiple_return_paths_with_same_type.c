// vybe-test: c/functions_advanced/function_can_have_multiple_return_paths_with_same_type
// origin: languages/c/tests/c/test_functions_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
double clamp_unit(double x) { if (x < 0.0) return 0.0; if (x > 1.0) return 1.0; return x; }
int main() {
const char *__w[] = {"0.0 0.5 1.0\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%.1f %.1f %.1f\n", clamp_unit(-1.0), clamp_unit(0.5), clamp_unit(2.0));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

