// vybe-test: c/cover_headers_misc/pp_file_macro
// origin: languages/c/tests/c/test_cover_headers_misc.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
return __FILE__ != 0;
}

