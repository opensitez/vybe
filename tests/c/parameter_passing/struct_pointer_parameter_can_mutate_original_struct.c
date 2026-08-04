// vybe-test: c/parameter_passing/struct_pointer_parameter_can_mutate_original_struct
// origin: languages/c/tests/c/test_parameter_passing.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Pair { int a; int b; }; void change(struct Pair *pair) { pair->a = 9; }
int main() {
const char *__w[] = {"9\n"};
int __n = 1, __i = 0;
struct Pair pair = {1,2}; change(&pair); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", pair.a);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

