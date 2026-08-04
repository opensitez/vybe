// vybe-test: c/wide_chars/wchar_string_literal_compiles
// origin: languages/c/tests/c/test_wide_chars.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <wchar.h>
int main() {const char *__w[] = {"1\n"};
int __n = 1, __i = 0;

    const wchar_t *s = L"wide string";
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", s != NULL ? 1 : 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

