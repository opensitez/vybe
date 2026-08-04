// vybe-test: c/stdio_snprintf_buffer/snprintf_zero_size_no_write
// origin: languages/c/tests/c/test_stdio_snprintf_buffer.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"z\n"};
int __n = 1, __i = 0;
char b[4]={'z','z','z','z'}; snprintf(b,0,"%s","x"); { char __t[512]; snprintf(__t, sizeof(__t), "%c\n", b[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

