// vybe-test: c/string_tokenize_span_search/strtok_trailing_delimiter_ignored
// origin: languages/c/tests/c/test_string_tokenize_span_search.rs
#include <assert.h>
#include <stdio.h>
#include <string.h>
char s[] = "x,y,";
int main() {
const char *__w[] = {"x\n", "y\n", "1\n"};
int __n = 3, __i = 0;
char *tok=strtok(s, ","); { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", tok);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } tok=strtok(NULL, ","); { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", tok);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } tok=strtok(NULL, ","); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", tok==NULL);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

