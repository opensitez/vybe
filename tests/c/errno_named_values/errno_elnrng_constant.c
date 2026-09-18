// vybe-test: c/errno_named_values/errno_elnrng_constant
// origin: languages/c/tests/c/test_errno_named_values.rs
// vybe-test-mode: compile
#include <errno.h>
#ifndef ELNRNG
#define ELNRNG 48
#endif
int main() {
return ELNRNG;
}

