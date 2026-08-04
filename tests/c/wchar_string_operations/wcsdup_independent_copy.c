// vybe-test: c/wchar_string_operations/wcsdup_independent_copy
// origin: languages/c/tests/c/test_wchar_string_operations.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <wchar.h>
#include <stdlib.h>
int main() {
const char *__w[] = {"N\n"};
int __n = 1, __i = 0;
wchar_t *p = wcsdup(L"ok"); p[0]=L'N'; { char __t[512]; snprintf(__t, sizeof(__t), "%lc\n", p[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } free(p); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

