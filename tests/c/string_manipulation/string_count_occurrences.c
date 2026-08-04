// vybe-test: c/string_manipulation/string_count_occurrences
// origin: languages/c/tests/c/test_string_manipulation.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;

char s[] = "banana";
int count = 0;
for (int i = 0; s[i]; i++) if (s[i] == 'a') count++;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", count);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

