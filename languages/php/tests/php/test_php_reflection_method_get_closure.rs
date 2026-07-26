use super::helpers::run_prints;

#[test]
fn test_reflection_method_get_closure_instance() {
    assert_eq!(
        run_prints(
            r#"<?php
class Calculator {
    public function add(int $a, int $b): int { return $a + $b; }
}
$calc = new Calculator();
$rm = new ReflectionMethod($calc, 'add');
$closure = $rm->getClosure($calc);
echo $closure(10, 20), "\n";
"#
        ),
        vec!["30"]
    );
}

#[test]
fn test_reflection_method_get_closure_static() {
    assert_eq!(
        run_prints(
            r#"<?php
class MathUtil {
    public static function square(int $n): int { return $n * $n; }
}
$rm = new ReflectionMethod(MathUtil::class, 'square');
$closure = $rm->getClosure();
echo $closure(7), "\n";
"#
        ),
        vec!["49"]
    );
}
