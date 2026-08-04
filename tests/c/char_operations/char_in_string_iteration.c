// vybe-test: c/char_operations/char_in_string_iteration
// origin: languages/c/tests/c/test_char_operations.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"a\n", "b\n", "c\n"};
int __n = 3, __i = 0;

char s[] = "abc";
for (int i = 0; s[i] != '\0'; i++) {
    { char __t[512]; snprintf(__t, sizeof(__t), "%c\n", s[i]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
}
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

