// vybe-test: c/cover_headers_misc/pp_ifndef_else
// origin: languages/c/tests/c/test_cover_headers_misc.rs
// vybe-test-mode: compile
#include <stdio.h>
#ifndef Z9
#define Z9
int x=2;
#endif
int main() {
return x;
}

