// vybe-test: c/parameter_passing/pointer_parameter_can_write_character_buffer
// origin: languages/c/tests/c/test_parameter_passing.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
void set_first(char *text) { text[0] = 'X'; }
int main() {
const char *__w[] = {"Xbc\n"};
int __n = 1, __i = 0;
char text[] = "abc"; set_first(text); { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", text);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

