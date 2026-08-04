// vybe-test: c/cover_string_h/strdup_compile
// origin: languages/c/tests/c/test_cover_string_h.rs
// vybe-test-mode: compile
#include <string.h>
#include <stdlib.h>
int main() {
char *s = strdup("x"); free(s); return 0;
}

