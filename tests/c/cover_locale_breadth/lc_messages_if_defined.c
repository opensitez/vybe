// vybe-test: c/cover_locale_breadth/lc_messages_if_defined
// origin: languages/c/tests/c/test_cover_locale_breadth.rs
// vybe-test-mode: compile
#include <locale.h>
#ifdef LC_MESSAGES
int use_lc_messages(void){return LC_MESSAGES;}
#endif
int main() {
return 0;
}

