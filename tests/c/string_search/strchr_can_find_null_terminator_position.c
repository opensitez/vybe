// vybe-test: c/string_search/strchr_can_find_null_terminator_position
// origin: languages/c/tests/c/test_string_search.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"term\n"};
int __n = 1, __i = 0;
if (strchr("abc", '\0') != NULL) { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "term");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } else { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "bad");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

