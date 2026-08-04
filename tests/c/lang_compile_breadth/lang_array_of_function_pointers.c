// vybe-test: c/lang_compile_breadth/lang_array_of_function_pointers
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
int f(int x){return x;} int (*tab[1])(int)={f};
int main() {
return tab[0](2);
}

