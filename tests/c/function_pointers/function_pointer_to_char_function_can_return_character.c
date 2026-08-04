// vybe-test: c/function_pointers/function_pointer_to_char_function_can_return_character
// origin: languages/c/tests/c/test_function_pointers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
char next_letter(char c) { return c + 1; }
int main() {
const char *__w[] = {"b\n"};
int __n = 1, __i = 0;
char (*fp)(char) = next_letter;
{ char __t[512]; snprintf(__t, sizeof(__t), "%c\n", fp('a'));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

