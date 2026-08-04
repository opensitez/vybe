// vybe-test: c/conditional_logic/dangling_else_attaches_to_nearest
// origin: languages/c/tests/c/test_conditional_logic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"small\n"};
int __n = 1, __i = 0;
int x = 1;
if (x > 0) if (x > 10) { char __t[512]; snprintf(__t, sizeof(__t), "big\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } else { char __t[512]; snprintf(__t, sizeof(__t), "small\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

