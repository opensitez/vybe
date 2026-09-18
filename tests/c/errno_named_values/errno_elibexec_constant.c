// vybe-test: c/errno_named_values/errno_elibexec_constant
// origin: languages/c/tests/c/test_errno_named_values.rs
// vybe-test-mode: compile
#include <errno.h>
#ifndef ELIBEXEC
#define ELIBEXEC 83
#endif
int main() {
return ELIBEXEC;
}

