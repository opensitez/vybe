// vybe-test: c/errno_named_values/errno_ecomm_constant
// origin: languages/c/tests/c/test_errno_named_values.rs
// vybe-test-mode: compile
#include <errno.h>
#ifndef ECOMM
#define ECOMM 70
#endif
int main() {
return ECOMM;
}

