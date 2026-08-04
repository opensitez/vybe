// vybe-test: c/lang_run_breadth4/anonymous_union_in_struct
// origin: languages/c/tests/c/test_lang_run_breadth4.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct S{ union { int i; char c; }; };
int main() {
const char *__w[] = {"A\n"};
int __n = 1, __i = 0;
struct S s; s.i=65; { char __t[512]; snprintf(__t, sizeof(__t), "%c\n", s.c);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

