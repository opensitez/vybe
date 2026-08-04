// vybe-test: c/cover_headers_misc/alignas_type
// origin: languages/c/tests/c/test_cover_headers_misc.rs
// vybe-test-mode: compile
#include <stdalign.h>
alignas(double) char buf[16];
int main() {
return sizeof(buf);
}

