// vybe-test: c/cover_breadth_batch_a/math_isinf
// origin: languages/c/tests/c/test_cover_breadth_batch_a.rs
// vybe-test-mode: compile
#include <math.h>
int main() {
return isinf(1.0/0.0);
}

