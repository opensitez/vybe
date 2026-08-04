// vybe-test: c/lang_semantics_batch/alignof_expression
// origin: languages/c/tests/c/test_lang_semantics_batch.rs
// vybe-test-mode: compile
#include <stdalign.h>
int main() {
return (int)alignof(double);
}

