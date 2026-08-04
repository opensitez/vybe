// vybe-test: c/lang_compile_breadth/lang_switch_goto_case
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
switch(1){case 1: goto L; L: return 1; } return 0;
}

