// vybe-test: c/char_operations/char_to_int_cast
// origin: languages/c/tests/c/test_char_operations.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"90\n"};
int __n = 1, __i = 0;
char c = 'Z';
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (int)c);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

