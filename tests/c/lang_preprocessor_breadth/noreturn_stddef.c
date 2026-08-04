// vybe-test: c/lang_preprocessor_breadth/noreturn_stddef
// origin: languages/c/tests/c/test_lang_preprocessor_breadth.rs
// vybe-test-mode: compile
#include <stdnoreturn.h>
_Noreturn void halt(void); void halt(void){for(;;){}}
int main() {
return 0;
}

