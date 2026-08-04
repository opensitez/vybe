// vybe-test: c/cover_string_h/strncpy_compile
// origin: languages/c/tests/c/test_cover_string_h.rs
// vybe-test-mode: compile
#include <string.h>
int main() {
char d[4]; strncpy(d, "abc", 3); return 0;
}

