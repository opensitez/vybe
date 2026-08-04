// vybe-test: c/flexible_array/flexible_array_int_data
// origin: languages/c/tests/c/test_flexible_array.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <stdlib.h>
struct IntVec { int n; int data[]; };
int main() {const char *__w[] = {"10\n", "20\n", "30\n"};
int __n = 3, __i = 0;

    struct IntVec *v = (struct IntVec*)malloc(sizeof(struct IntVec) + 3 * sizeof(int));
    v->n = 3;
    v->data[0] = 10; v->data[1] = 20; v->data[2] = 30;
    for (int i = 0; i < v->n; i++) { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", v->data[i]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    free(v);
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

