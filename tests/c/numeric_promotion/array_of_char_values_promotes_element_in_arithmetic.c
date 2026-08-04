// vybe-test: c/numeric_promotion/array_of_char_values_promotes_element_in_arithmetic
// origin: languages/c/tests/c/test_numeric_promotion.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"98\n"};
int __n = 1, __i = 0;
char letters[2] = {'a', 'b'}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", letters[0] + 1);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

