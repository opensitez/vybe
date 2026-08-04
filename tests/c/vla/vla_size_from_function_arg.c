// vybe-test: c/vla/vla_size_from_function_arg
// origin: languages/c/tests/c/test_vla.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"1\n", "2\n", "3\n"};
static int __n = 3, __i = 0;

#include <stdio.h>
void fill(int n) {
    int arr[n];
    for (int i = 0; i < n; i++) arr[i] = i + 1;
    for (int i = 0; i < n; i++) { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", arr[i]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
}
int main() {
    fill(3);
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

