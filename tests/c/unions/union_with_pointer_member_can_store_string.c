// vybe-test: c/unions/union_with_pointer_member_can_store_string
// origin: languages/c/tests/c/test_unions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
union View { char *text; int i; };
int main() {
const char *__w[] = {"vybe\n"};
int __n = 1, __i = 0;
union View view; view.text = "vybe"; { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", view.text);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

