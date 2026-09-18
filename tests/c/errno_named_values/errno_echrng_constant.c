// vybe-test: c/errno_named_values/errno_echrng_constant
// origin: languages/c/tests/c/test_errno_named_values.rs
// vybe-test-mode: compile
#include <errno.h>
#ifndef ECHRNG
#define ECHRNG 44
#endif
int main() {
return ECHRNG;
}

