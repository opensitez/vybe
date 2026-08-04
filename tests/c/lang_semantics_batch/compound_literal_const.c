// vybe-test: c/lang_semantics_batch/compound_literal_const
// origin: languages/c/tests/c/test_lang_semantics_batch.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
int *p=(int[]){1,2}; return p[0];
}

