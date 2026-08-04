// vybe-test: c/lang_compile_breadth/lang_do_while_continue
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
int i=0; do{ if(++i<2) continue; break; } while(1); return i;
}

