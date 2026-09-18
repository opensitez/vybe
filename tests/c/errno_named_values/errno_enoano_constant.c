// vybe-test: c/errno_named_values/errno_enoano_constant
// origin: languages/c/tests/c/test_errno_named_values.rs
// vybe-test-mode: compile
#include <errno.h>
#ifndef ENOANO
#define ENOANO 55
#endif
int main() {
return ENOANO;
}

