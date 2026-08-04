// vybe-test: c/lang_preprocessor_breadth/static_assert_msg
// origin: languages/c/tests/c/test_lang_preprocessor_breadth.rs
// vybe-test-mode: compile
#include <assert.h>
_Static_assert(sizeof(int)>=4, "int");
int main() {
return 0;
}

