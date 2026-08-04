// vybe-test: c/pointers_arithmetic/pointer_loop_can_emit_each_character
// origin: languages/c/tests/c/test_pointers_arithmetic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
char text[] = "go"; char *p = text;
int main() {
const char *__w[] = {"g\n", "o\n"};
int __n = 2, __i = 0;
while (*p) { { char __t[512]; snprintf(__t, sizeof(__t), "%c\n", *p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } p++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

