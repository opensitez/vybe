// vybe-test: c/lang_compile_breadth/lang_static_assert_type
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <assert.h>
_Static_assert(sizeof(char)==1,"");
int main() {
return 0;
}

