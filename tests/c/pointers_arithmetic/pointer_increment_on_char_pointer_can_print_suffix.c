// vybe-test: c/pointers_arithmetic/pointer_increment_on_char_pointer_can_print_suffix
// origin: languages/c/tests/c/test_pointers_arithmetic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
char text[] = "world"; char *p = text;
int main() {
const char *__w[] = {"orld\n"};
int __n = 1, __i = 0;
p++;
{ char __t[512]; snprintf(__t, sizeof(__t), "%s\n", p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

