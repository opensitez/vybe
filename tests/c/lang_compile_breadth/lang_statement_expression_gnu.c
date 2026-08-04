// vybe-test: c/lang_compile_breadth/lang_statement_expression_gnu
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
int x=({int y=2; y+1;}); return x;
}

