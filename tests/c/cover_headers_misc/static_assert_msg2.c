// vybe-test: c/cover_headers_misc/static_assert_msg2
// origin: languages/c/tests/c/test_cover_headers_misc.rs
// vybe-test-mode: compile
#include <assert.h>
_Static_assert(sizeof(void*)>=4,"ptr");
int main() {
return 0;
}

