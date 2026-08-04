// vybe-test: c/c_if_else_dangling/if_else_dangling_switch
// origin: languages/c/tests/c/test_c_if_else_dangling.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 int x=1; if (1) switch(0) case 0: if (1) x=2; else x=3; { char __t[512]; snprintf(__t, sizeof(__t), "%d", x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

