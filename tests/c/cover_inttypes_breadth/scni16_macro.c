// vybe-test: c/cover_inttypes_breadth/scni16_macro
// origin: languages/c/tests/c/test_cover_inttypes_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#include <inttypes.h>
int main() {
int16_t v=0; sscanf("1", "%" SCNi16, &v); return v;
}

