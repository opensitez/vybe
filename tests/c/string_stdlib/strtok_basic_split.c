// vybe-test: c/string_stdlib/strtok_basic_split
// origin: languages/c/tests/c/test_string_stdlib.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"a\n", "b\n", "c\n"};
int __n = 3, __i = 0;

char s[] = "a:b:c";
char *tok = strtok(s, ":");
while (tok) {
    { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", tok);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    tok = strtok(NULL, ":");
}
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

