// vybe-test: c/enums/enum_with_values
// origin: languages/c/tests/c/test_enums.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
enum Status { OK = 200, NOT_FOUND = 404, ERROR = 500 };
int main() {const char *__w[] = {"200\n", "404\n", "500\n"};
int __n = 3, __i = 0;

    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", OK);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", NOT_FOUND);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", ERROR);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

