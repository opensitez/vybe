// vybe-test: c/string_tokenize_span_search/strtok_modifies_buffer_in_place
// origin: languages/c/tests/c/test_string_tokenize_span_search.rs
#include <assert.h>
#include <stdio.h>
#include <string.h>
char s[] = "p:q";
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
strtok(s, ":"); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", s[1]=='\0');
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

