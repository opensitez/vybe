// vybe-test: c/stdlib_misc/calloc_char_buffer_can_hold_string_then_free
// origin: languages/c/tests/c/test_stdlib_misc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"hi\n"};
int __n = 1, __i = 0;
char *p = (char *)calloc(8, sizeof(char)); p[0] = 'h'; p[1] = 'i'; { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } free(p); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

