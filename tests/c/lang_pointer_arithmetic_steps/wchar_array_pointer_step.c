// vybe-test: c/lang_pointer_arithmetic_steps/wchar_array_pointer_step
// origin: languages/c/tests/c/test_lang_pointer_arithmetic_steps.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <wchar.h>
int main() {
const char *__w[] = {"b\n"};
int __n = 1, __i = 0;
wchar_t s[] = L"ab"; wchar_t *p=s; p++; { char __t[512]; snprintf(__t, sizeof(__t), "%lc\n", *p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

