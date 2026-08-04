// vybe-test: c/unions/union_can_be_returned_from_function
// origin: languages/c/tests/c/test_unions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
union Data { int i; char c; }; union Data make_data(int x) { union Data data; data.i = x; return data; }
int main() {
const char *__w[] = {"66\n"};
int __n = 1, __i = 0;
union Data data = make_data(66); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", data.i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

