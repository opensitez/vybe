// vybe-test: c/lang_preprocessor_breadth/thread_local_macro
// origin: languages/c/tests/c/test_lang_preprocessor_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
_Thread_local int tls;
int main() {
tls=1; return tls;
}

