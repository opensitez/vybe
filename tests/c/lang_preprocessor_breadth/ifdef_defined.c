// vybe-test: c/lang_preprocessor_breadth/ifdef_defined
// origin: languages/c/tests/c/test_lang_preprocessor_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#define F
#ifdef F
int a=1;
#endif
int main() {
return a;
}

