// vybe-test: c/string_manipulation/string_is_palindrome
// origin: languages/c/tests/c/test_string_manipulation.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;

char s[] = "racecar";
int len = strlen(s), ok = 1;
for (int i = 0; i < len/2; i++) if (s[i] != s[len-1-i]) { ok = 0; break; }
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", ok);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

