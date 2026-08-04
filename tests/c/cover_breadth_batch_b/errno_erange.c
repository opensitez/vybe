// vybe-test: c/cover_breadth_batch_b/errno_erange
// origin: languages/c/tests/c/test_cover_breadth_batch_b.rs
// vybe-test-mode: compile
#include <errno.h>
int main() {
return ERANGE;
}

