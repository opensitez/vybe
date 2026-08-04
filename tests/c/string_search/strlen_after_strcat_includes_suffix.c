// vybe-test: c/string_search/strlen_after_strcat_includes_suffix
// origin: languages/c/tests/c/test_string_search.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
char left[32] = "moon";
int main() {
const char *__w[] = {"9\n"};
int __n = 1, __i = 0;
strcat(left, "light");
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", strlen(left));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

