// vybe-test: c/global_variables/global_modified_by_function
// origin: languages/c/tests/c/test_global_variables.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int counter = 0;
void increment() { counter++; }
int main() {const char *__w[] = {"3\n"};
int __n = 1, __i = 0;

    increment();
    increment();
    increment();
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", counter);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

