// vybe-test: c/enum_advanced2/enum_bitmask_flags
// origin: languages/c/tests/c/test_enum_advanced2.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef enum { NONE=0, READ=1, WRITE=2, EXEC=4 } Perms;
int main() {
const char *__w[] = {"1\n", "0\n"};
int __n = 2, __i = 0;
Perms p = READ | WRITE;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (p & READ) != 0 ? 1 : 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (p & EXEC) != 0 ? 1 : 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

