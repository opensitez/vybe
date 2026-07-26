use super::helpers::run_prints;

#[test]
fn test_reflection_enum_get_backing_type_string() {
    assert_eq!(
        run_prints(
            r#"<?php
enum StringEnum: string {
    case Alpha = 'a';
}
$re = new ReflectionEnum(StringEnum::class);
$type = $re->getBackingType();
echo $type instanceof ReflectionNamedType ? $type->getName() : 'none', "\n";
"#
        ),
        vec!["string"]
    );
}

#[test]
fn test_reflection_enum_get_backing_type_int() {
    assert_eq!(
        run_prints(
            r#"<?php
enum IntEnum: int {
    case One = 1;
}
$re = new ReflectionEnum(IntEnum::class);
$type = $re->getBackingType();
echo $type instanceof ReflectionNamedType ? $type->getName() : 'none', "\n";
"#
        ),
        vec!["int"]
    );
}
