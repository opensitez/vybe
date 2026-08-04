// vybe-test: c/lang_semantics_batch/void_expr_compile
// origin: languages/c/tests/c/test_lang_semantics_batch.rs
// vybe-test-mode: compile
#include <stdio.h>
void noop(void){}
int main() {
noop(); return 0;
}

