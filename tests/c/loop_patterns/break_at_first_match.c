// vybe-test: c/loop_patterns/break_at_first_match
// origin: languages/c/tests/c/test_loop_patterns.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"2\n"};
int __n = 1, __i = 0;

int arr[] = {10, 20, 30, 40, 50};
int target = 30, idx = -1;
for (int i = 0; i < 5; i++) {
    if (arr[i] == target) { idx = i; break; }
}
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", idx);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

