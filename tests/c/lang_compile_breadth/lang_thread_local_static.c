// vybe-test: c/lang_compile_breadth/lang_thread_local_static
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
_Thread_local static int tls;
int main() {
tls=1; return tls;
}

