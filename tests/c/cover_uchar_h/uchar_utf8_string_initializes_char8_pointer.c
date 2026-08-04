// vybe-test: c/cover_uchar_h/uchar_utf8_string_initializes_char8_pointer
// origin: languages/c/tests/c/test_cover_uchar_h.rs
// vybe-test-mode: compile
#include <uchar.h>
int main() {
const char8_t *s = u8"ok"; return s[0];
}

