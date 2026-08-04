// vybe-test: c/string_locale/strxfrm_sortable_result
// origin: languages/c/tests/c/test_string_locale.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;

char a[64], b[64];
strxfrm(a, "abc", sizeof(a));
strxfrm(b, "abd", sizeof(b));
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", strcmp(a, b) < 0 ? 1 : 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

