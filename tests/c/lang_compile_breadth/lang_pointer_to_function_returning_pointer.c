// vybe-test: c/lang_compile_breadth/lang_pointer_to_function_returning_pointer
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
int *f(void){ static int x=1; return &x; }
int main() {
int *(*fp)(void)=f; return *fp();
}

