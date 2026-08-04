// vybe-test: c/cover_inttypes_breadth/prix64_macro
// origin: languages/c/tests/c/test_cover_inttypes_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#include <inttypes.h>
int main() {
uint64_t v=15; printf("%" PRIx64, v); return 0;
}

