// vybe-test: c/string_memory_search_ops/memrchr_n_one
// origin: languages/c/tests/c/test_string_memory_search_ops.rs
#include <assert.h>
#include <stdio.h>
#include <string.h>
int main() {
const char *__w[] = {"y\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%c\n", *memrchr("yz", 'y', 1));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

