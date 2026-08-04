// vybe-test: c/vla/vla_basic_declaration
// origin: languages/c/tests/c/test_vla.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int main() {const char *__w[] = {"0 4 8\n"};
int __n = 1, __i = 0;

    int n = 5;
    int arr[n];
    for (int i = 0; i < n; i++) arr[i] = i * 2;
    { char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d\n", arr[0], arr[2], arr[4]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

