// vybe-test: c/lang_compile_breadth/lang_label_at_end
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
goto L; L: return 0;
}

