// vybe-test: c/lang_preprocessor_breadth/elif_branch
// origin: languages/c/tests/c/test_lang_preprocessor_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#if 0
int d=1;
#elif 1
int d=2;
#endif
int main() {
return d;
}

