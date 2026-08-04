// vybe-test: c/string_manipulation/string_remove_char
// origin: languages/c/tests/c/test_string_manipulation.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"heo word\n"};
int __n = 1, __i = 0;

char s[] = "hello world";
int w = 0;
for (int r = 0; s[r]; r++) if (s[r] != 'l') s[w++] = s[r];
s[w] = '\0';
{ char __t[512]; snprintf(__t, sizeof(__t), "%s\n", s);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

