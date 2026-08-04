// vybe-test: c/cover_uchar_h/uchar_mbstate_t_declares
// origin: languages/c/tests/c/test_cover_uchar_h.rs
// vybe-test-mode: compile
#include <uchar.h>
int main() {
mbstate_t st; (void)st; return 0;
}

