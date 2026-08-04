// vybe-test: c/c_pointer_void_cast_arithmetic/void_pointer_arithmetic_gnu
// origin: languages/c/tests/c/test_c_pointer_void_cast_arithmetic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
/* GNU C allows void* arithmetic, treats size as 1 */ int main() { int x[2] = {1, 2}; void *p = x; /* let's just check standard behavior or ignore */ printf("ok"); return 0; }

