// vybe-test: c/cover_inttypes_breadth/prix32_macro
// origin: languages/c/tests/c/test_cover_inttypes_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#include <inttypes.h>
int main() {
uint32_t v=15; printf("%" PRIx32, v); return 0;
}

