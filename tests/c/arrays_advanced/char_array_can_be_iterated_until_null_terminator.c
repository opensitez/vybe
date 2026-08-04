// vybe-test: c/arrays_advanced/char_array_can_be_iterated_until_null_terminator
// origin: languages/c/tests/c/test_arrays_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
char text[] = "go";
int main() {
const char *__w[] = {"g\n", "o\n"};
int __n = 2, __i = 0;
int i = 0;
while (text[i]) { { char __t[512]; snprintf(__t, sizeof(__t), "%c\n", text[i]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

