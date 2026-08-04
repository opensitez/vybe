// vybe-test: c/multidim_arrays/two_dim_char_array_strings
// origin: languages/c/tests/c/test_multidim_arrays.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"one two three\n"};
int __n = 1, __i = 0;
char words[3][6] = {"one", "two", "three"};
{ char __t[512]; snprintf(__t, sizeof(__t), "%s %s %s\n", words[0], words[1], words[2]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

