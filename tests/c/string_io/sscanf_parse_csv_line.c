// vybe-test: c/string_io/sscanf_parse_csv_line
// origin: languages/c/tests/c/test_string_io.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"Alice 30 9.5\n"};
int __n = 1, __i = 0;

char name[16]; int age; float score;
sscanf("Alice,30,9.5", "%15[^,],%d,%f", name, &age, &score);
{ char __t[512]; snprintf(__t, sizeof(__t), "%s %d %.1f\n", name, age, score);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

