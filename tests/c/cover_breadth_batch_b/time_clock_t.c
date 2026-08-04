// vybe-test: c/cover_breadth_batch_b/time_clock_t
// origin: languages/c/tests/c/test_cover_breadth_batch_b.rs
// vybe-test-mode: compile
#include <time.h>
int main() {
clock_t c=clock(); return c>=0;
}

