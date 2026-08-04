// vybe-test: c/lang_control_scope/internal_static_function
// origin: languages/c/tests/c/test_lang_control_scope.rs
// vybe-test-mode: compile
#include <stdio.h>
static int helper(void){return 1;}
int main() {
return helper();
}

