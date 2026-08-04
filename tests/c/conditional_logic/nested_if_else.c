// vybe-test: c/conditional_logic/nested_if_else
// origin: languages/c/tests/c/test_conditional_logic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"both pos\n"};
int __n = 1, __i = 0;

int a = 1, b = 2;
if (a > 0) {
    if (b > 0) { char __t[512]; snprintf(__t, sizeof(__t), "both pos\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    else { char __t[512]; snprintf(__t, sizeof(__t), "a pos\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
} else {
    { char __t[512]; snprintf(__t, sizeof(__t), "a neg\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
}
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

