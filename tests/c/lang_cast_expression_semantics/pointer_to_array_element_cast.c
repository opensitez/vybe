// vybe-test: c/lang_cast_expression_semantics/pointer_to_array_element_cast
// origin: languages/c/tests/c/test_lang_cast_expression_semantics.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"20\n"};
int __n = 1, __i = 0;
int arr[3] = {10, 20, 30}; int *mid = (int *)&arr[1]; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", *mid);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

