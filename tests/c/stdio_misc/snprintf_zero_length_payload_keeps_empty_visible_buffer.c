// vybe-test: c/stdio_misc/snprintf_zero_length_payload_keeps_empty_visible_buffer
// origin: languages/c/tests/c/test_stdio_misc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
char buf[8] = "seed";
int main() {
const char *__w[] = {"\n"};
int __n = 1, __i = 0;
snprintf(buf, 1, "%s", "abc");
{ char __t[512]; snprintf(__t, sizeof(__t), "%s\n", buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

