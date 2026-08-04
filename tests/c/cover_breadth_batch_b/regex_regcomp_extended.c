// vybe-test: c/cover_breadth_batch_b/regex_regcomp_extended
// origin: languages/c/tests/c/test_cover_breadth_batch_b.rs
// vybe-test-mode: compile
#include <regex.h>
int main() {
regex_t r; regcomp(&r,"a",REG_EXTENDED); regfree(&r); return 0;
}

