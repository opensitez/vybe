// vybe-test: c/lang_run_breadth3/struct_member_pointer
// origin: languages/c/tests/c/test_lang_run_breadth3.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct S{int n;};
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;
struct S s={2}; int *p=&s.n; *p=3; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", s.n);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

