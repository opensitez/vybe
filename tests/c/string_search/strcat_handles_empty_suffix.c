// vybe-test: c/string_search/strcat_handles_empty_suffix
// origin: languages/c/tests/c/test_string_search.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
char left[32] = "hello";
int main() {
const char *__w[] = {"hello\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%s\n", strcat(left, ""));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

