// vybe-test: c/lang_compile_breadth/lang_case_range_extension
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
switch(2){ case 1 ... 3: return 1; } return 0;
}

