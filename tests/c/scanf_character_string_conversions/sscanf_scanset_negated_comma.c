// vybe-test: c/scanf_character_string_conversions/sscanf_scanset_negated_comma
// origin: languages/c/tests/c/test_scanf_character_string_conversions.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"field\n"};
int __n = 1, __i = 0;
char buf[16]; sscanf("field,rest", "%[^,]", buf); { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

