// Python cmath — complex math: phase, polar, rect, sqrt, exp, log, trig
use super::helpers::run_python;

#[test]
fn test_cmath_sqrt_negative() {
    let script = r#"
import cmath
result = cmath.sqrt(-1)
print(result.real)
print(round(result.imag, 10))
"#;
    assert_eq!(run_python(script), vec!["0.0", "1.0"]);
}

#[test]
fn test_cmath_phase() {
    let script = r#"
import cmath, math
z = complex(1, 1)
p = cmath.phase(z)
print(round(p, 5))
print(round(p, 5) == round(math.pi / 4, 5))
"#;
    assert_eq!(run_python(script), vec!["0.7854", "True"]);
}

#[test]
fn test_cmath_polar() {
    let script = r#"
import cmath
r, phi = cmath.polar(complex(3, 4))
print(round(r, 5))
"#;
    assert_eq!(run_python(script), vec!["5.0"]);
}

#[test]
fn test_cmath_rect() {
    let script = r#"
import cmath, math
z = cmath.rect(1, math.pi)
print(round(z.real, 5))
print(round(z.imag, 5))
"#;
    assert_eq!(run_python(script), vec!["-1.0", "0.0"]);
}

#[test]
fn test_cmath_exp() {
    let script = r#"
import cmath, math
z = cmath.exp(complex(0, math.pi))
print(round(z.real, 5))
print(round(z.imag, 5))
"#;
    assert_eq!(run_python(script), vec!["-1.0", "0.0"]);
}

#[test]
fn test_cmath_log() {
    let script = r#"
import cmath
z = cmath.log(complex(1, 0))
print(round(z.real, 5))
print(round(z.imag, 5))
"#;
    assert_eq!(run_python(script), vec!["0.0", "0.0"]);
}

#[test]
fn test_cmath_isfinite_isinf_isnan() {
    let script = r#"
import cmath
print(cmath.isfinite(1 + 2j))
print(cmath.isinf(complex(float('inf'), 0)))
print(cmath.isnan(complex(float('nan'), 0)))
"#;
    assert_eq!(run_python(script), vec!["True", "True", "True"]);
}

#[test]
fn test_cmath_isclose() {
    let script = r#"
import cmath
print(cmath.isclose(1+1j, 1+1j))
print(cmath.isclose(1+1j, 1+2j))
"#;
    assert_eq!(run_python(script), vec!["True", "False"]);
}

#[test]
fn test_cmath_sin_cos() {
    let script = r#"
import cmath
z = cmath.sin(0)
print(round(z.real, 5))
z2 = cmath.cos(0)
print(round(z2.real, 5))
"#;
    assert_eq!(run_python(script), vec!["0.0", "1.0"]);
}
