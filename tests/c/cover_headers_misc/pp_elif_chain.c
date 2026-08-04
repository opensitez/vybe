// vybe-test: c/cover_headers_misc/pp_elif_chain
// origin: languages/c/tests/c/test_cover_headers_misc.rs
// vybe-test-mode: compile
#include <stdio.h>
#if 0
int x=1;
#elif 1
int x=2;
#else
int x=3;
#endif
int main() {
return x;
}

