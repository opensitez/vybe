// vybe-test: c/errno_named_values/errno_elibacc_constant
// origin: languages/c/tests/c/test_errno_named_values.rs
// vybe-test-mode: compile
#include <errno.h>
#ifndef ELIBACC
#define ELIBACC 79
#endif
int main() {
return ELIBACC;
}

