// vybe-test: c/errno_named_values/errno_exfull_constant
// origin: languages/c/tests/c/test_errno_named_values.rs
// vybe-test-mode: compile
#include <errno.h>
#ifndef EXFULL
#define EXFULL 54
#endif
int main() {
return EXFULL;
}

