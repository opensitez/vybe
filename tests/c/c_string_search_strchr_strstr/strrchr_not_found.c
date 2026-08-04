// vybe-test: c/c_string_search_strchr_strstr/strrchr_not_found
// origin: languages/c/tests/c/test_c_string_search_strchr_strstr.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <string.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 char *s = "hello"; { char __t[512]; snprintf(__t, sizeof(__t), "%d", strrchr(s, 'x') == NULL);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

