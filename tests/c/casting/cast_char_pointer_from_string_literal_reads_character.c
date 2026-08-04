// vybe-test: c/casting/cast_char_pointer_from_string_literal_reads_character
// origin: languages/c/tests/c/test_casting.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
char *p = (char *)"vybe";
int main() {
const char *__w[] = {"y\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%c\n", p[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

