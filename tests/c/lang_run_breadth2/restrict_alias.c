// vybe-test: c/lang_run_breadth2/restrict_alias
// origin: languages/c/tests/c/test_lang_run_breadth2.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
void addp(restrict int *a, restrict int *b){*a+=*b;}
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;
int x=1,y=2; addp(&x,&y); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

