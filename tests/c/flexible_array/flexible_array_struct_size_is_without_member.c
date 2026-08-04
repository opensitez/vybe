// vybe-test: c/flexible_array/flexible_array_struct_size_is_without_member
// origin: languages/c/tests/c/test_flexible_array.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
struct Flex { int n; double data[]; };
int main() {const char *__w[] = {"8\n"};
int __n = 1, __i = 0;

    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (int)sizeof(struct Flex));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

