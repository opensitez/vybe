// vybe-test: c/cover_headers_misc/alignof_max
// origin: languages/c/tests/c/test_cover_headers_misc.rs
// vybe-test-mode: compile
#include <stdalign.h>
int main() {
return alignof(max_align_t);
}

