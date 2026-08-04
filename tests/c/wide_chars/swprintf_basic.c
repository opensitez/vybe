// vybe-test: c/wide_chars/swprintf_basic
// origin: languages/c/tests/c/test_wide_chars.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <wchar.h>
int main() {const char *__w[] = {"1\n"};
int __n = 1, __i = 0;

    wchar_t buf[32];
    swprintf(buf, 32, L"val=%d", 99);
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (int)wcslen(buf) > 0 ? 1 : 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

