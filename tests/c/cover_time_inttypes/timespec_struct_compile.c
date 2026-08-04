// vybe-test: c/cover_time_inttypes/timespec_struct_compile
// origin: languages/c/tests/c/test_cover_time_inttypes.rs
// vybe-test-mode: compile
#include <time.h>
int main() {
struct timespec ts = {0,0}; return ts.tv_sec;
}

