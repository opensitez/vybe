// vybe-test: c/cover_wchar_misc/fenv_feraiseexcept
// origin: languages/c/tests/c/test_cover_wchar_misc.rs
// vybe-test-mode: compile
#include <fenv.h>
int main() {
feraiseexcept(FE_INVALID); return 0;
}

