// vybe-test: c/pointer_to_array/string_pointer_array_traverse
// origin: languages/c/tests/c/test_pointer_to_array.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"one\n", "two\n", "three\n"};
int __n = 3, __i = 0;

const char *words[] = {"one","two","three",NULL};
const char **p = words;
while (*p) { { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", *p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } p++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

