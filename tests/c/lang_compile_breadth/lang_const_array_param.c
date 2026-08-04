// vybe-test: c/lang_compile_breadth/lang_const_array_param
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
int len(const int *a){return a[0];}
int main() {
int a[]={5}; return len(a);
}

