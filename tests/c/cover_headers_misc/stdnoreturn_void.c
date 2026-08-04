// vybe-test: c/cover_headers_misc/stdnoreturn_void
// origin: languages/c/tests/c/test_cover_headers_misc.rs
// vybe-test-mode: compile
#include <stdnoreturn.h>
_Noreturn void h(void); void h(void){for(;;){}}
int main() {
return 0;
}

