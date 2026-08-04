// vybe-test: c/cover_headers_misc/pp_error_if_zero
// origin: languages/c/tests/c/test_cover_headers_misc.rs
// vybe-test-mode: compile
#include <stdio.h>
#if 1
int x=1;
#endif
int main() {
return x;
}

