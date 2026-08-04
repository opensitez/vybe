// vybe-test: c/enum_advanced2/enum_negative_values
// origin: languages/c/tests/c/test_enum_advanced2.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
enum Err { ERR_OK=0, ERR_FAIL=-1, ERR_FATAL=-99 };
int main() {
const char *__w[] = {"-1 -99\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", ERR_FAIL, ERR_FATAL);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

