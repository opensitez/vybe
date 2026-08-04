// vybe-test: c/complex_initializers/static_array_with_string_elements
// origin: languages/c/tests/c/test_complex_initializers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"Jan Dec\n"};
int __n = 1, __i = 0;

static const char *MONTHS[] = {"Jan","Feb","Mar","Apr","May","Jun",
    "Jul","Aug","Sep","Oct","Nov","Dec"};
{ char __t[512]; snprintf(__t, sizeof(__t), "%s %s\n", MONTHS[0], MONTHS[11]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

