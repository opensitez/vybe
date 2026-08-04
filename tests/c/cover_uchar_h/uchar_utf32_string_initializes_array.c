// vybe-test: c/cover_uchar_h/uchar_utf32_string_initializes_array
// origin: languages/c/tests/c/test_cover_uchar_h.rs
// vybe-test-mode: compile
#include <uchar.h>
int main() {
const char32_t *s = U"ok"; return s[1];
}

