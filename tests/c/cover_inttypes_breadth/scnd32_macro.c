// vybe-test: c/cover_inttypes_breadth/scnd32_macro
// origin: languages/c/tests/c/test_cover_inttypes_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#include <inttypes.h>
int main() {
int32_t v=0; sscanf("1", "%" SCNd32, &v); return v;
}

