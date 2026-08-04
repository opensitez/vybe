// vybe-test: c/lang_compile_breadth/lang_vla_parameter
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
void f(int n, int a[n]){(void)a;}
int main() {
f(1,(int[]){1}); return 0;
}

