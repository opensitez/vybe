// vybe-test: c/cover_complex_limits/complex_catan
// origin: languages/c/tests/c/test_cover_complex_limits.rs
// vybe-test-mode: compile
#include <complex.h>
int main() {
double complex z = catan(0); return (int)creal(z);
}

