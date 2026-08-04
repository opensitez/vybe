// vybe-test: c/cover_headers_misc/complex_h_macro_i
// origin: languages/c/tests/c/test_cover_headers_misc.rs
// vybe-test-mode: compile
#include <complex.h>
int main() {
double complex z=I; return cimag(z)==1;
}

