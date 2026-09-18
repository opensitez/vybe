// vybe-test: c/errno_named_values/errno_ebfont_constant
// origin: languages/c/tests/c/test_errno_named_values.rs
// vybe-test-mode: compile
#include <errno.h>
#ifndef EBFONT
#define EBFONT 59
#endif
int main() {
return EBFONT;
}

