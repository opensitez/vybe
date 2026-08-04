// vybe-test: c/lang_pointer_arithmetic_steps/pointer_arithmetic_with_sizeof_element
// origin: languages/c/tests/c/test_lang_pointer_arithmetic_steps.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"99\n"};
int __n = 1, __i = 0;
int a[2]={42,99}; char *bp=(char*)a; bp+=sizeof(int); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", *(int*)bp);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

