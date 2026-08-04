// vybe-test: c/lang_pointer_arithmetic_steps/array_of_pointers_increment_selects_next
// origin: languages/c/tests/c/test_lang_pointer_arithmetic_steps.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"2\n"};
int __n = 1, __i = 0;
int x=1,y=2,z=3; int *a[3]={&x,&y,&z}; int **p=a; p++; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", **p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

