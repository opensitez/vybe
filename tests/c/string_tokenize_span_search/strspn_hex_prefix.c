// vybe-test: c/string_tokenize_span_search/strspn_hex_prefix
// origin: languages/c/tests/c/test_string_tokenize_span_search.rs
#include <assert.h>
#include <stdio.h>
#include <string.h>
int main() {
const char *__w[] = {"6\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (int)strspn("0xFF00", "0123456789ABCDEFabcdefx"));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

