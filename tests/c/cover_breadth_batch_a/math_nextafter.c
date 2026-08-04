// vybe-test: c/cover_breadth_batch_a/math_nextafter
// origin: languages/c/tests/c/test_cover_breadth_batch_a.rs
// vybe-test-mode: compile
#include <math.h>
int main() {
return nextafter(1.0,2.0) > 1.0;
}

