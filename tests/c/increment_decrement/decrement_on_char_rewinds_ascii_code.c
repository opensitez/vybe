// vybe-test: c/increment_decrement/decrement_on_char_rewinds_ascii_code
// origin: languages/c/tests/c/test_increment_decrement.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
char c = 'b';
int main() {
const char *__w[] = {"a\n"};
int __n = 1, __i = 0;
--c;
{ char __t[512]; snprintf(__t, sizeof(__t), "%c\n", c);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

