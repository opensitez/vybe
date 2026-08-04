// vybe-test: c/cover_time_inttypes/nanosleep_compile
// origin: languages/c/tests/c/test_cover_time_inttypes.rs
// vybe-test-mode: compile
#include <time.h>
int main() {
struct timespec r={0,0}, req={0,0}; nanosleep(&req,&r); return 0;
}

