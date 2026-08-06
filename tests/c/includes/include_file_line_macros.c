// vybe-test: c/includes/include_file_line_macros
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"1 1"};
int __n = 1, __i = 0;
int has_name = strstr(__FILE__, "include_file_line_macros") != NULL;
int line_ok = __LINE__ == 9;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d", has_name, line_ok);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}
