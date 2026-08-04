// vybe-test: c/c_struct_bitfields_signedness/bitfield_sign_signed_int
// origin: languages/c/tests/c/test_c_struct_bitfields_signedness.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct S { signed int a:3; }; int main() { struct S s; s.a = 7; /* 7 is 111, in 3-bit signed it's -1 */ printf("%d", s.a); return 0; }

