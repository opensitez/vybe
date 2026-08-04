// vybe-test: c/time_strftime_directives/strftime_percent_n_newline_compile
// origin: languages/c/tests/c/test_time_strftime_directives.rs
// vybe-test-mode: compile
#include <stdio.h>
#include <time.h>
int main() {
struct tm t={0}; char b[4]; strftime(b,sizeof(b),"%n",&t); return 0;
}

