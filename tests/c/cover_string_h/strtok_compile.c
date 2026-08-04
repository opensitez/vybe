// vybe-test: c/cover_string_h/strtok_compile
// origin: languages/c/tests/c/test_cover_string_h.rs
// vybe-test-mode: compile
#include <string.h>
int main() {
char s[]="a:b"; strtok(s,":"); return 0;
}

