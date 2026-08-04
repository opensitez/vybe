// vybe-test: c/data_structures/dynamic_array_append
// origin: languages/c/tests/c/test_data_structures.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

typedef struct { int *data; int len; int cap; } Vec;
void vec_push(Vec *v, int x) {
    if (v->len >= v->cap) {
        v->cap = v->cap ? v->cap * 2 : 4;
        v->data = (int*)realloc(v->data, v->cap * sizeof(int));
    }
    v->data[v->len++] = x;
}
int main() {
const char *__w[] = {"0\n", "10\n", "20\n", "30\n", "40\n"};
int __n = 5, __i = 0;

Vec v = {NULL, 0, 0};
for (int i = 0; i < 5; i++) vec_push(&v, i * 10);
for (int i = 0; i < v.len; i++) { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", v.data[i]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
free(v.data);
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

