use super::helpers::run_prints;

#[test]
fn test_reflection_function_has_return_type() {
    assert_eq!(
        run_prints(
            r#"<?php
function calculate(int $a): int { return $a * 2; }
$rf = new ReflectionFunction('calculate');
echo $rf->hasReturnType() ? 'has_return_type' : 'none', "\n";
"#
        ),
        vec!["has_return_type"]
    );
}

#[test]
fn test_reflection_function_nullable_return_type() {
    assert_eq!(
        run_prints(
            r#"<?php
function find_user(int $id): ?string { return $id === 1 ? 'Alice' : null; }
$rf = new ReflectionFunction('find_user');
$rt = $rf->getReturnType();
echo $rt->allowsNull() ? 'nullable_return' : 'strict', "\n";
"#
        ),
        vec!["nullable_return"]
    );
}

#[test]
fn test_reflection_function_closure_parameters() {
    assert_eq!(
        run_prints(
            r#"<?php
$closure = function(string $name, int $age = 18): string { return "$name:$age"; };
$rf = new ReflectionFunction($closure);
echo $rf->getNumberOfParameters() . ':' . $rf->getNumberOfRequiredParameters(), "\n";
"#
        ),
        vec!["2:1"]
    );
}
