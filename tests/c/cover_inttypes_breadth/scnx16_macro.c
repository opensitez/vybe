// vybe-test: c/cover_inttypes_breadth/scnx16_macro
// origin: languages/c/tests/c/test_cover_inttypes_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#include <inttypes.h>
int main() {
uint16_t v=0; sscanf("f", "%" SCNx16, &v); return v;
}

