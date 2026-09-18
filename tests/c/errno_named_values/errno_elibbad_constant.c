// vybe-test: c/errno_named_values/errno_elibbad_constant
// origin: languages/c/tests/c/test_errno_named_values.rs
// vybe-test-mode: compile
#include <errno.h>
#ifndef ELIBBAD
#define ELIBBAD 80
#endif
int main() {
return ELIBBAD;
}

