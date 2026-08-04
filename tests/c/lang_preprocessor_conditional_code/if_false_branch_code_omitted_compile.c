// vybe-test: c/lang_preprocessor_conditional_code/if_false_branch_code_omitted_compile
// origin: languages/c/tests/c/test_lang_preprocessor_conditional_code.rs
// vybe-test-mode: compile
#include <stdio.h>
#if 0
void dead(void);
#endif
int x=1;
int main() {
return x;
}

