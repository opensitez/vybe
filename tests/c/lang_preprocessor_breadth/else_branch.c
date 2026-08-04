// vybe-test: c/lang_preprocessor_breadth/else_branch
// origin: languages/c/tests/c/test_lang_preprocessor_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#if 0
int e=1;
#else
int e=2;
#endif
int main() {
return e;
}

