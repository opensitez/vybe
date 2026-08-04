// vybe-test: c/cover_headers_misc/pp_time_macro
// origin: languages/c/tests/c/test_cover_headers_misc.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
return __TIME__[0];
}

