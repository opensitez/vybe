// vybe-test: c/memory_functions/memmove_overlapping
// origin: languages/c/tests/c/test_memory_functions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"aabcd\n"};
int __n = 1, __i = 0;

char buf[] = "abcde";
memmove(buf + 1, buf, 4);
{ char __t[512]; snprintf(__t, sizeof(__t), "%c%c%c%c%c\n", buf[0], buf[1], buf[2], buf[3], buf[4]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

