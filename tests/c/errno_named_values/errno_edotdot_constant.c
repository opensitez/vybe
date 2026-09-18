// vybe-test: c/errno_named_values/errno_edotdot_constant
// origin: languages/c/tests/c/test_errno_named_values.rs
// vybe-test-mode: compile
#include <errno.h>
#ifndef EDOTDOT
#define EDOTDOT 73
#endif
int main() {
return EDOTDOT;
}

