// vybe-test: c/cover_inttypes_breadth/scnu16_macro
// origin: languages/c/tests/c/test_cover_inttypes_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#include <inttypes.h>
int main() {
uint16_t v=0; sscanf("1", "%" SCNu16, &v); return v;
}

