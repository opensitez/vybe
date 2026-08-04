// vybe-test: c/pointer_advanced/dynamic_array_via_malloc
// origin: languages/c/tests/c/test_pointer_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"0 1 4 9\n"};
int __n = 1, __i = 0;

int n = 4;
int *arr = (int*)malloc(n * sizeof(int));
for (int i = 0; i < n; i++) arr[i] = i * i;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d %d\n", arr[0], arr[1], arr[2], arr[3]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
free(arr);
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

