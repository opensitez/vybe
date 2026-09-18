// vybe-test: c/errno_named_values/errno_ebadrqc_constant
// origin: languages/c/tests/c/test_errno_named_values.rs
// vybe-test-mode: compile
#include <errno.h>
#ifndef EBADRQC
#define EBADRQC 56
#endif
int main() {
return EBADRQC;
}

