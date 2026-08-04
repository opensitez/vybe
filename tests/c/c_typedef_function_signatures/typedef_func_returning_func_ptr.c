// vybe-test: c/c_typedef_function_signatures/typedef_func_returning_func_ptr
// origin: languages/c/tests/c/test_c_typedef_function_signatures.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"A"};
static int __n = 1, __i = 0;
typedef void (*Action)(void); typedef Action (*GetAction)(void); void a() { { char __t[512]; snprintf(__t, sizeof(__t), "A");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } Action get() { return a; } int main() { GetAction g = get; g()(); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

