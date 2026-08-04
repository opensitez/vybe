// vybe-test: c/cover_wchar_misc/assert_static_compile
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <assert.h>
int main() {
static_assert(1,""); return 0;
}

