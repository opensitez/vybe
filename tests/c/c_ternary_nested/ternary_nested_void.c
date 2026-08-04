// vybe-test: c/c_ternary_nested/ternary_nested_void
// origin: languages/c/tests/c/test_c_ternary_nested.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"G"};
static int __n = 1, __i = 0;
void f() { { char __t[512]; snprintf(__t, sizeof(__t), "F");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } void g() { { char __t[512]; snprintf(__t, sizeof(__t), "G");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } int main() { 1 ? 0 ? f() : g() : f(); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

