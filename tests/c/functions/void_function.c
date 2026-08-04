// vybe-test: c/functions/void_function
// origin: languages/c/tests/c/test_functions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"Hello Alice\n", "Hello Bob\n"};
static int __n = 2, __i = 0;

#include <stdio.h>
void greet(char *name) {
    { char __t[512]; snprintf(__t, sizeof(__t), "Hello %s\n", name);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
}
int main() {
    greet("Alice");
    greet("Bob");
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

