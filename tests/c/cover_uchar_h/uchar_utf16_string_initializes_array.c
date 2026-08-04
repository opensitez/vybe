// vybe-test: c/cover_uchar_h/uchar_utf16_string_initializes_array
// origin: languages/c/tests/c/test_cover_uchar_h.rs
// vybe-test-mode: compile
#include <uchar.h>
int main() {
const char16_t *s = u"ok"; return s[1];
}

