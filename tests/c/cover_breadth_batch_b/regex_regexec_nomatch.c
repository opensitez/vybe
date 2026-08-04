// vybe-test: c/cover_breadth_batch_b/regex_regexec_nomatch
// origin: languages/c/tests/c/test_cover_breadth_batch_b.rs
// vybe-test-mode: compile
#include <regex.h>
int main() {
regex_t r; regcomp(&r,"z",0); regexec(&r,"a",0,0,0); regfree(&r); return 0;
}

