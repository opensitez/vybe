// vybe-test: c/cover_breadth_batch_a/string_strerror_r
// origin: languages/c/tests/c/test_cover_breadth_batch_a.rs
// vybe-test-mode: compile
#include <string.h>
int main() {
char b[64]; return strerror_r(EINVAL,b,64) != 0;
}

