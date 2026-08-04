// vybe-test: c/lang_control_scope/nested_block_compile
// origin: languages/c/tests/c/test_lang_control_scope.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
{ { int x=1; } } return 0;
}

