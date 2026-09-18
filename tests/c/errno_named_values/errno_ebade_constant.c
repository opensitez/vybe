// vybe-test: c/errno_named_values/errno_ebade_constant
// origin: languages/c/tests/c/test_errno_named_values.rs
// vybe-test-mode: compile
#include <errno.h>
#ifndef EBADE
#define EBADE 52
#endif
int main() {
return EBADE;
}

