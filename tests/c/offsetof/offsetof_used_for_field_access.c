// vybe-test: c/offsetof/offsetof_used_for_field_access
// origin: languages/c/tests/c/test_offsetof.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct S { int x; int y; };
int main() {
const char *__w[] = {"20\n"};
int __n = 1, __i = 0;

struct S s = {10, 20};
char *base = (char*)&s;
int *yp = (int*)(base + offsetof(struct S, y));
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", *yp);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

