// vybe-test: c/cover_wchar_misc/fenv_fetestexcept_div
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <fenv.h>
int main() {
return fetestexcept(FE_DIVBYZERO);
}

