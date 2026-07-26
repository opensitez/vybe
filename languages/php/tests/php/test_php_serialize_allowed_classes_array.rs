use super::helpers::run_prints;

#[test]
fn test_unserialize_allowed_classes_array_whitelist() {
    assert_eq!(
        run_prints(
            r#"<?php
class AllowedDto {
    public string $data = 'safe';
}
class ForbiddenDto {
    public string $data = 'unsafe';
}

$s1 = serialize(new AllowedDto());
$s2 = serialize(new ForbiddenDto());

$obj1 = unserialize($s1, ['allowed_classes' => [AllowedDto::class]]);
$obj2 = unserialize($s2, ['allowed_classes' => [AllowedDto::class]]);

echo ($obj1 instanceof AllowedDto ? 'allowed_ok' : 'err') . '|' . ($obj2 instanceof __PHP_Incomplete_Class ? 'incomplete_ok' : 'err'), "\n";
"#
        ),
        vec!["allowed_ok|incomplete_ok"]
    );
}

#[test]
fn test_unserialize_allowed_classes_false() {
    assert_eq!(
        run_prints(
            r#"<?php
class PayloadDto {
    public int $id = 1;
}
$s = serialize(new PayloadDto());
$obj = unserialize($s, ['allowed_classes' => false]);
echo ($obj instanceof __PHP_Incomplete_Class) ? 'disallowed_all_ok' : 'err', "\n";
"#
        ),
        vec!["disallowed_all_ok"]
    );
}
