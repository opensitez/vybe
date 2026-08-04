// vybe-test: c/cover_inttypes_breadth/scnx32_macro
// origin: languages/c/tests/c/test_cover_inttypes_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#include <inttypes.h>
int main() {
uint32_t v=0; sscanf("f", "%" SCNx32, &v); return v;
}

