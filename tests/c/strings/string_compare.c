// vybe-test: c/strings/string_compare
// origin: languages/c/tests/c/test_strings.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <string.h>
int main() {const char *__w[] = {"equal\n", "not equal\n"};
int __n = 2, __i = 0;

    char *a = "hello";
    char *b = "hello";
    char *c = "world";
    if (strcmp(a, b) == 0) { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "equal");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    else { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "not equal");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (strcmp(a, c) == 0) { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "equal");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    else { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "not equal");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

