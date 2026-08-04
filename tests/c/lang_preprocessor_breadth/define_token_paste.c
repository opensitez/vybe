// vybe-test: c/lang_preprocessor_breadth/define_token_paste
// origin: languages/c/tests/c/test_lang_preprocessor_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#define CONCAT(a,b) a##b
int xy = 1;
int main() {
return CONCAT(x,y);
}

