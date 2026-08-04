// vybe-test: c/lang_run_breadth3/while_continue
// origin: languages/c/tests/c/test_lang_run_breadth3.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"4\n"};
int __n = 1, __i = 0;
int i=0,s=0; while(i++<3){ if(i==2) continue; s+=i;} { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", s);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

