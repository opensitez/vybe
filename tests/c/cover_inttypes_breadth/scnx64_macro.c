// vybe-test: c/cover_inttypes_breadth/scnx64_macro
// origin: languages/c/tests/c/test_cover_inttypes_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#include <inttypes.h>
int main() {
uint64_t v=0; sscanf("f", "%" SCNx64, &v); return v;
}

