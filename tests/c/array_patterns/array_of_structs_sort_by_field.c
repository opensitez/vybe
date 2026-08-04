// vybe-test: c/array_patterns/array_of_structs_sort_by_field
// origin: languages/c/tests/c/test_array_patterns.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Item { int id; int val; };
int main() {
const char *__w[] = {"10 20 30\n"};
int __n = 1, __i = 0;

struct Item items[3] = {{1,30},{2,10},{3,20}};
for (int i = 0; i < 2; i++)
    for (int j = 0; j < 2-i; j++)
        if (items[j].val > items[j+1].val) {
            struct Item t = items[j]; items[j]=items[j+1]; items[j+1]=t;
        }
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d\n", items[0].val, items[1].val, items[2].val);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

