// vybe-test: c/cover_headers_misc/pp_defined_and
// origin: languages/c/tests/c/test_cover_headers_misc.rs
// vybe-test-mode: compile
#include <stdio.h>
#if defined(__STDC__) && 1
int x=1;
#endif
int main() {
return x;
}

