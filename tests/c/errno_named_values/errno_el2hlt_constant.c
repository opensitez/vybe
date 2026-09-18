// vybe-test: c/errno_named_values/errno_el2hlt_constant
// origin: languages/c/tests/c/test_errno_named_values.rs
// vybe-test-mode: compile
#include <errno.h>
#ifndef EL2HLT
#define EL2HLT 51
#endif
int main() {
return EL2HLT;
}

