// vybe-test: c/complex_typedefs/typedef_pointer_to_struct
// origin: languages/c/tests/c/test_complex_typedefs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Node { int val; };
typedef struct Node *NodePtr;
int main() {
const char *__w[] = {"55\n"};
int __n = 1, __i = 0;
struct Node n; n.val = 55;
NodePtr p = &n;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", p->val);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

