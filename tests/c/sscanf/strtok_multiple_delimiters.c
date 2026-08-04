// vybe-test: c/sscanf/strtok_multiple_delimiters
// origin: languages/c/tests/c/test_sscanf.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"one\n", "two\n", "three\n"};
int __n = 3, __i = 0;

char s[] = "one two\tthree";
char *tok = strtok(s, " \t");
while (tok != NULL) {
    { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", tok);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    tok = strtok(NULL, " \t");
}
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

