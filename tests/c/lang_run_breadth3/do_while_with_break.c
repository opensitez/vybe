// vybe-test: c/lang_run_breadth3/do_while_with_break
// origin: languages/c/tests/c/test_lang_run_breadth3.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"2\n"};
int __n = 1, __i = 0;
int i=0; do{ if(++i==2) break; } while(1); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

