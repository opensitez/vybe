// vybe-test: c/struct_methods/struct_method_modifies_by_pointer
// origin: languages/c/tests/c/test_struct_methods.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

typedef struct { int count; } Counter;
void counter_increment(Counter *c) { c->count++; }
void counter_reset(Counter *c) { c->count = 0; }
int main() {
const char *__w[] = {"3\n", "0\n"};
int __n = 2, __i = 0;
Counter c = {0};
counter_increment(&c);
counter_increment(&c);
counter_increment(&c);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", c.count);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
counter_reset(&c);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", c.count);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

