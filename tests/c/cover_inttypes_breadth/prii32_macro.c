// vybe-test: c/cover_inttypes_breadth/prii32_macro
// origin: languages/c/tests/c/test_cover_inttypes_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#include <inttypes.h>
int main() {
int32_t v=1; printf("%" PRIi32, v); return 0;
}

