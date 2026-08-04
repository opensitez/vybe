// vybe-test: c/lang_sizeof_alignof_expressions/alignof_max_align_t
// origin: languages/c/tests/c/test_lang_sizeof_alignof_expressions.rs
// vybe-test-mode: compile
#include <stdalign.h>
int main() {
return (int)alignof(max_align_t);
}

