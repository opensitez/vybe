// vybe-test: c/functions/recursive_fibonacci
// origin: languages/c/tests/c/test_functions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int fib(int n) {
    if (n <= 1) return n;
    return fib(n - 1) + fib(n - 2);
}
int main() {const char *__w[] = {"0\n", "1\n", "13\n"};
int __n = 3, __i = 0;

    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", fib(0));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", fib(1));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", fib(7));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

