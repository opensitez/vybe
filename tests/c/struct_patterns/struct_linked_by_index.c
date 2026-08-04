// vybe-test: c/struct_patterns/struct_linked_by_index
// origin: languages/c/tests/c/test_struct_patterns.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Item { int val; int next; };
int main() {
const char *__w[] = {"10\n", "20\n", "30\n"};
int __n = 3, __i = 0;

struct Item list[3] = {{10, 1}, {20, 2}, {30, -1}};
int idx = 0;
while (idx >= 0) {
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", list[idx].val);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    idx = list[idx].next;
}
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

