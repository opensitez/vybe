// vybe-test: c/memory_functions/memchr_finds_byte
// origin: languages/c/tests/c/test_memory_functions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"2\n"};
int __n = 1, __i = 0;

char buf[] = "hello";
char *p = (char*)memchr(buf, 'l', 5);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (int)(p - buf));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

