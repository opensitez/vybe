// vybe-test: c/cover_breadth_batch_a/regex_regerror
// origin: languages/c/tests/c/test_cover_breadth_batch_a.rs
// vybe-test-mode: compile
#include <regex.h>
int main() {
regex_t r; char e[64]; regcomp(&r,"[",REG_EXTENDED); regerror(0,&r,e,64); regfree(&r); return 0;
}

