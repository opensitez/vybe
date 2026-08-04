// vybe-test: c/string_manipulation/string_to_uppercase_manual
// origin: languages/c/tests/c/test_string_manipulation.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"HELLO\n"};
int __n = 1, __i = 0;

char s[] = "hello";
for (int i = 0; s[i]; i++) if (s[i] >= 'a' && s[i] <= 'z') s[i] -= 32;
{ char __t[512]; snprintf(__t, sizeof(__t), "%s\n", s);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

