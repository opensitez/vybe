use super::helpers::*;

// C99 complex number support via <complex.h>
#[test]
fn complex_real_and_imaginary_parts() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <complex.h>
int main() {
    double complex z = 3.0 + 4.0 * I;
    printf("%.1f %.1f\n", creal(z), cimag(z));
    return 0;
}
"#,
        &["3.0 4.0"],
    );
}

#[test]
fn complex_addition() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <complex.h>
int main() {
    double complex a = 1.0 + 2.0 * I;
    double complex b = 3.0 + 4.0 * I;
    double complex c = a + b;
    printf("%.1f %.1f\n", creal(c), cimag(c));
    return 0;
}
"#,
        &["4.0 6.0"],
    );
}

#[test]
fn complex_multiplication() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <complex.h>
int main() {
    double complex a = 1.0 + 2.0 * I;
    double complex b = 3.0 + 4.0 * I;
    double complex c = a * b;
    printf("%.1f %.1f\n", creal(c), cimag(c));
    return 0;
}
"#,
        &["-5.0 10.0"],
    );
}

#[test]
fn complex_absolute_value() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <complex.h>
int main() {
    double complex z = 3.0 + 4.0 * I;
    printf("%.1f\n", cabs(z));
    return 0;
}
"#,
        &["5.0"],
    );
}

#[test]
fn complex_conjugate() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <complex.h>
int main() {
    double complex z = 3.0 + 4.0 * I;
    double complex c = conj(z);
    printf("%.1f %.1f\n", creal(c), cimag(c));
    return 0;
}
"#,
        &["3.0 -4.0"],
    );
}

#[test]
fn complex_argument() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <complex.h>
#include <math.h>
int main() {
    double complex z = 0.0 + 1.0 * I;
    double arg = carg(z);
    printf("%.4f\n", arg);
    return 0;
}
"#,
        &["1.5708"],
    );
}

#[test]
fn float_complex_type() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <complex.h>
int main() {
    float complex z = 2.0f + 3.0f * I;
    printf("%.1f\n", crealf(z));
    return 0;
}
"#,
        &["2.0"],
    );
}
