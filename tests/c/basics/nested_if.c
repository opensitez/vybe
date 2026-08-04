// vybe-test: c/basics/nested_if
// origin: languages/c/tests/c/test_basics.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int main() {const char *__w[] = {"big positive\n"};
int __n = 1, __i = 0;

    int x = 5;
    if (x > 0) {
        if (x > 3) {
            { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "big positive");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
        } else {
            { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "small positive");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
        }
    } else {
        { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "negative");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

