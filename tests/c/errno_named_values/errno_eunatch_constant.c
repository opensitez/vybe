// vybe-test: c/errno_named_values/errno_eunatch_constant
// origin: languages/c/tests/c/test_errno_named_values.rs
// vybe-test-mode: compile
#include <errno.h>
#ifndef EUNATCH
#define EUNATCH 49
#endif
int main() {
return EUNATCH;
}

