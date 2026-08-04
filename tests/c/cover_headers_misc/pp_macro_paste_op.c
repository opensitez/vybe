// vybe-test: c/cover_headers_misc/pp_macro_paste_op
// origin: languages/c/tests/c/test_cover_headers_misc.rs
// vybe-test-mode: compile
#include <stdio.h>
#define C(a,b) a##b
int xy=1;
int main() {
return C(x,y);
}

