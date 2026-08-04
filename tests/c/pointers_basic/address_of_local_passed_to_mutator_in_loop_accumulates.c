// vybe-test: c/pointers_basic/address_of_local_passed_to_mutator_in_loop_accumulates
// origin: languages/c/tests/c/test_pointers_basic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
void bump(int *p) { (*p)++; }
int main() {
const char *__w[] = {"4\n"};
int __n = 1, __i = 0;
int n = 0;
for (int i = 0; i < 4; i++) bump(&n);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", n);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

