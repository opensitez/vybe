// vybe-test: c/lang_array_decay_parameters/char_array_param_as_pointer_string
// origin: languages/c/tests/c/test_lang_array_decay_parameters.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int len(char a[]){ int n=0; while(a[n]) n++; return n; }
int main() {
const char *__w[] = {"4\n"};
int __n = 1, __i = 0;
char s[]="four"; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", len(s));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

