// vybe-test: c/string_manipulation/string_word_count
// origin: languages/c/tests/c/test_string_manipulation.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"4\n"};
int __n = 1, __i = 0;

char s[] = "one two three four";
int words = 0, in_word = 0;
for (int i = 0; s[i]; i++) {
    if (s[i] != ' ') { if (!in_word) { words++; in_word = 1; } }
    else in_word = 0;
}
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", words);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

