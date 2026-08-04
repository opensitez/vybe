// vybe-test: c/cover_headers_misc/ctype_isblank_tab
// origin: languages/c/tests/c/test_cover_headers_misc.rs
// vybe-test-mode: compile
#include <ctype.h>
int main() {
return isblank('\t');
}

