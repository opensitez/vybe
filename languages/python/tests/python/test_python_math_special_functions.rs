use super::helpers::run_python;

#[test]
fn test_python_math_factorial_comb_perm() {
    let src = r#"
import math
print(math.factorial(5))
print(math.comb(5, 2))
print(math.perm(5, 3))
"#;
    assert_eq!(run_python(src), vec!["120", "10", "60"]);
}

#[test]
fn test_python_math_gcd_lcm() {
    let src = r#"
import math
print(math.gcd(12, 18))
print(math.lcm(4, 6, 8))
"#;
    assert_eq!(run_python(src), vec!["6", "24"]);
}
