// vybe-test: c/pointers_arithmetic/pointer_difference_on_char_array_counts_characters
// origin: languages/c/tests/c/test_pointers_arithmetic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
char text[] = "hello";
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (int)(&text[4] - &text[1]));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

