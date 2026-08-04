// vybe-test: c/designated_init/struct_designated_init_nested
// origin: languages/c/tests/c/test_designated_init.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Inner { int a; int b; }; struct Outer { struct Inner in; int c; };
int main() {
const char *__w[] = {"1 2 3\n"};
int __n = 1, __i = 0;
struct Outer o = {.in={.a=1, .b=2}, .c=3};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d\n", o.in.a, o.in.b, o.c);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

