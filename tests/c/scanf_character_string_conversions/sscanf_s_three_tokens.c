// vybe-test: c/scanf_character_string_conversions/sscanf_s_three_tokens
// origin: languages/c/tests/c/test_scanf_character_string_conversions.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"red/green/blue\n"};
int __n = 1, __i = 0;
char a[6],b[6],c[6]; sscanf("red green blue", "%s %s %s", a,b,c); { char __t[512]; snprintf(__t, sizeof(__t), "%s/%s/%s\n", a,b,c);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

