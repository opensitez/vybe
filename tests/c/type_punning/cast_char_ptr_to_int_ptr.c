// vybe-test: c/type_punning/cast_char_ptr_to_int_ptr
// origin: languages/c/tests/c/test_type_punning.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
int main() {const char *__w[] = {"1\n"};
int __n = 1, __i = 0;

    int x = 0x01020304;
    char *p = (char*)&x;
    int same = *p != 0;
    { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", same);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

