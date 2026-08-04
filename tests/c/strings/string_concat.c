// vybe-test: c/strings/string_concat
// origin: languages/c/tests/c/test_strings.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <string.h>
int main() {const char *__w[] = {"hello world\n"};
int __n = 1, __i = 0;

    char *a = "hello";
    char *b = " world";
    char *c = strcat(a, b);
    { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", c);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

