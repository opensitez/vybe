// vybe-test: c/unions/union_of_struct_and_int_can_read_struct_member
// origin: languages/c/tests/c/test_unions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Pair { int a; int b; }; union Mixed { struct Pair pair; int i; };
int main() {
const char *__w[] = {"3 4\n"};
int __n = 1, __i = 0;
union Mixed mixed; mixed.pair.a = 3; mixed.pair.b = 4; { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", mixed.pair.a, mixed.pair.b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

