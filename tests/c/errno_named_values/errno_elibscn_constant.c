// vybe-test: c/errno_named_values/errno_elibscn_constant
// origin: languages/c/tests/c/test_errno_named_values.rs
// vybe-test-mode: compile
#include <errno.h>
#ifndef ELIBSCN
#define ELIBSCN 81
#endif
int main() {
return ELIBSCN;
}

