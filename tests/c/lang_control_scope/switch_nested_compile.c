// vybe-test: c/lang_control_scope/switch_nested_compile
// origin: languages/c/tests/c/test_lang_control_scope.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
switch(1){case 1: switch(2){case 2: break;} break;} return 0;
}

