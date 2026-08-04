// vybe-test: c/typedef_patterns/typedef_pointer_opaque
// origin: languages/c/tests/c/test_typedef_patterns.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct _Handle { int id; };
typedef struct _Handle *Handle;
int main() {
const char *__w[] = {"99\n"};
int __n = 1, __i = 0;
struct _Handle h = {99};
Handle p = &h;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", p->id);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

