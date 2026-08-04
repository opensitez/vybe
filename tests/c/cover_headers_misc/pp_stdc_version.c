// vybe-test: c/cover_headers_misc/pp_stdc_version
// origin: languages/c/tests/c/test_cover_headers_misc.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
return __STDC_VERSION__ > 0;
}

