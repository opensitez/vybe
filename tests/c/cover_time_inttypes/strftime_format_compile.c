// vybe-test: c/cover_time_inttypes/strftime_format_compile
// origin: languages/c/tests/c/test_cover_time_inttypes.rs
// vybe-test-mode: compile
#include <time.h>
#include <stdio.h>
int main() {
char b[4]; struct tm t={0}; strftime(b,4,"%j",&t); return 0;
}

