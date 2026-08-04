// vybe-test: c/cover_inttypes_breadth/prix16_macro
// origin: languages/c/tests/c/test_cover_inttypes_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#include <inttypes.h>
int main() {
uint16_t v=15; printf("%" PRIx16, v); return 0;
}

