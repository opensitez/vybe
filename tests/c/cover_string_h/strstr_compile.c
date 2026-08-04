// vybe-test: c/cover_string_h/strstr_compile
// origin: languages/c/tests/c/test_cover_string_h.rs
// vybe-test-mode: compile
#include <string.h>
int main() {
return strstr("ab","b") != 0;
}

