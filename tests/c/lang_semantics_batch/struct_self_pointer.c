// vybe-test: c/lang_semantics_batch/struct_self_pointer
// origin: languages/c/tests/c/test_lang_semantics_batch.rs
// vybe-test-mode: compile
#include <stdio.h>
struct N { struct N *next; };
int main() {
struct N n={0}; return 0;
}

