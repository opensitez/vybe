// vybe-test: c/loop_patterns/for_loop_multiple_init_update
// origin: languages/c/tests/c/test_loop_patterns.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"0 10\n", "1 9\n", "2 8\n", "3 7\n", "4 6\n"};
int __n = 5, __i = 0;
int i, j;
for (i=0, j=10; i<5; i++, j--) { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", i, j);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

