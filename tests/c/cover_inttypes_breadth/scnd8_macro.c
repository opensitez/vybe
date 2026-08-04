// vybe-test: c/cover_inttypes_breadth/scnd8_macro
// origin: languages/c/tests/c/test_cover_inttypes_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#include <inttypes.h>
int main() {
int8_t v=0; sscanf("1", "%" SCNd8, &v); return v;
}

