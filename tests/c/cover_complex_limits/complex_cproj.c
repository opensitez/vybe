// vybe-test: c/cover_complex_limits/complex_cproj
// origin: languages/c/tests/c/test_cover_complex_limits.rs
// vybe-test-mode: compile
#include <complex.h>
int main() {
double complex z = cproj(1+2*I); return (int)creal(z);
}

