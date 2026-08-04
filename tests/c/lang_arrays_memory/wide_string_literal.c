// vybe-test: c/lang_arrays_memory/wide_string_literal
// origin: languages/c/tests/c/test_lang_arrays_memory.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <wchar.h>
int main() {
const char *__w[] = {"w\n"};
int __n = 1, __i = 0;
wchar_t *s = L"w"; { char __t[512]; snprintf(__t, sizeof(__t), "%lc\n", s[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

