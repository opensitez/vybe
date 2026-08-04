// vybe-test: c/lang_vla_stack_arrays/vla_char_buffer_string_build
// origin: languages/c/tests/c/test_lang_vla_stack_arrays.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"abc\n"};
int __n = 1, __i = 0;
int n=4; char s[n]; s[0]='a';s[1]='b';s[2]='c';s[3]='\0'; { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", s);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

