// vybe-test: c/printf_pointer_n_conversions/printf_n_after_percent_literal
// origin: languages/c/tests/c/test_printf_pointer_n_conversions.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <stddef.h>
int main() {
const char *__w[] = {"50%3"};
int __n = 1, __i = 0;
int n=0; { char __t[512]; snprintf(__t, sizeof(__t), "50%%%n", &n);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", n);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

