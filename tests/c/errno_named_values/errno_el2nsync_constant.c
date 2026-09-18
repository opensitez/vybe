// vybe-test: c/errno_named_values/errno_el2nsync_constant
// origin: languages/c/tests/c/test_errno_named_values.rs
// vybe-test-mode: compile
#include <errno.h>
#ifndef EL2NSYNC
#define EL2NSYNC 45
#endif
int main() {
return EL2NSYNC;
}

