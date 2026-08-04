// vybe-test: c/struct_patterns/struct_with_char_array
// origin: languages/c/tests/c/test_struct_patterns.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Name { char first[16]; char last[16]; };
int main() {
const char *__w[] = {"John Doe\n"};
int __n = 1, __i = 0;
struct Name n;
strcpy(n.first, "John");
strcpy(n.last, "Doe");
{ char __t[512]; snprintf(__t, sizeof(__t), "%s %s\n", n.first, n.last);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

