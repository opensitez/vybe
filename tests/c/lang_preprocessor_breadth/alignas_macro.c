// vybe-test: c/lang_preprocessor_breadth/alignas_macro
// origin: languages/c/tests/c/test_lang_preprocessor_breadth.rs
// vybe-test-mode: compile
#include <stdalign.h>
alignas(8) int x;
int main() {
return x;
}

