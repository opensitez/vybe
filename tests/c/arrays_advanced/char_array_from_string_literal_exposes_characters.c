// vybe-test: c/arrays_advanced/char_array_from_string_literal_exposes_characters
// origin: languages/c/tests/c/test_arrays_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
char text[] = "cat";
int main() {
const char *__w[] = {"c a t\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%c %c %c\n", text[0], text[1], text[2]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

