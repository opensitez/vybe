// vybe-test: c/c_char_classification_ascii_extended/ctype_isalnum_true
// origin: languages/c/tests/c/test_c_char_classification_ascii_extended.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <ctype.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "%d", isalnum('b') != 0 && isalnum('8') != 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

