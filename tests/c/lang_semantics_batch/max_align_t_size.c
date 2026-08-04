// vybe-test: c/lang_semantics_batch/max_align_t_size
// origin: languages/c/tests/c/test_lang_semantics_batch.rs
// vybe-test-mode: compile
#include <stddef.h>
int main() {
return (int)sizeof(max_align_t);
}

