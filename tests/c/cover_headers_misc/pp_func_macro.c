// vybe-test: c/cover_headers_misc/pp_func_macro
// origin: languages/c/tests/c/test_cover_headers_misc.rs
// vybe-test-mode: compile
#include <stdio.h>
int f(void){return __LINE__;}
int main() {
return f();
}

