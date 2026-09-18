// vybe-test: c/errno_named_values/errno_enopkg_constant
// origin: languages/c/tests/c/test_errno_named_values.rs
// vybe-test-mode: compile
#include <errno.h>
#ifndef ENOPKG
#define ENOPKG 65
#endif
int main() {
return ENOPKG;
}

