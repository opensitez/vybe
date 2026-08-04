// vybe-test: c/c_pointer_to_incomplete_type/pointer_incomplete_self_referential
// origin: languages/c/tests/c/test_c_pointer_to_incomplete_type.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Node { struct Node *next; }; int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 struct Node n; n.next = &n; { char __t[512]; snprintf(__t, sizeof(__t), "%d", n.next == &n);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

