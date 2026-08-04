// vybe-test: c/lang_preprocessor_breadth/macro_variadic
// origin: languages/c/tests/c/test_lang_preprocessor_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#define LOG(fmt, ...) printf(fmt, __VA_ARGS__)
int main() {
LOG("%d\n",1); return 0;
}

