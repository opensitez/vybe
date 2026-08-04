// vybe-test: c/lang_run_breadth2/designated_struct_init
// origin: languages/c/tests/c/test_lang_run_breadth2.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct S{int a,b;};
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;
struct S s={.b=3}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", s.b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

