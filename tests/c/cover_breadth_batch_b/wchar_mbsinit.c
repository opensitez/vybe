// vybe-test: c/cover_breadth_batch_b/wchar_mbsinit
// origin: languages/c/tests/c/test_cover_breadth_batch_b.rs
// vybe-test-mode: compile
#include <wchar.h>
int main() {
mbstate_t st; return mbsinit(&st);
}

