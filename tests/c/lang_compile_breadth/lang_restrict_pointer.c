// vybe-test: c/lang_compile_breadth/lang_restrict_pointer
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
void copy(int * restrict d, int * restrict s){*d=*s;}
int main() {
int a=1,b=2; copy(&a,&b); return a;
}

