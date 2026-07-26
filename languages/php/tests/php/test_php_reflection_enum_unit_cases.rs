use super::helpers::run_prints;

#[test]
fn test_reflection_enum_is_enum() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Status {
    case Pending;
    case Approved;
}

$re = new ReflectionEnum(Status::class);
echo $re->isEnum() ? 'is_enum' : 'not_enum', "\n";
"#
        ),
        vec!["is_enum"]
    );
}

#[test]
fn test_reflection_enum_get_cases() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Color {
    case Red;
    case Green;
    case Blue;
}

$re = new ReflectionEnum(Color::class);
$cases = $re->getCases();
echo count($cases) . ':' . $cases[0]->getName(), "\n";
"#
        ),
        vec!["3:Red"]
    );
}

#[test]
fn test_reflection_enum_backed_case_value() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Suit: string {
    case Hearts = 'H';
    case Diamonds = 'D';
}

$re = new ReflectionEnum(Suit::class);
echo $re->isBacked() ? 'backed' : 'pure', "\n";
$case = $re->getCase('Hearts');
echo $case->getBackingValue(), "\n";
"#
        ),
        vec!["backed", "H"]
    );
}

#[test]
fn test_reflection_enum_unit_case_get_enum() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Priority {
    case High;
}

$re = new ReflectionEnum(Priority::class);
$case = $re->getCase('High');
echo $case->getDeclaringClass()->getName(), "\n";
"#
        ),
        vec!["Priority"]
    );
}
