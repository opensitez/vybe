// vybe-test: c/cover_inttypes_breadth/prid64_macro
// origin: languages/c/tests/c/test_cover_inttypes_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#include <inttypes.h>
int main() {
int64_t v=1; printf("%" PRId64, v); return 0;
}

