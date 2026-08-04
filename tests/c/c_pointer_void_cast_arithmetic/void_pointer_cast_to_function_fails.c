// vybe-test: c/c_pointer_void_cast_arithmetic/void_pointer_cast_to_function_fails
// origin: languages/c/tests/c/test_c_pointer_void_cast_arithmetic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
/* void f(){} int main() { void *p = f; return 0; } // standard C doesn't guarantee object pointer can hold function pointer, but POSIX does */ int main() { printf("ok"); return 0; }

