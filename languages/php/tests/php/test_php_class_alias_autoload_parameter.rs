use super::helpers::run_prints;

#[test]
fn test_class_alias_autoload_true() {
    assert_eq!(
        run_prints(
            r#"<?php
class OriginalClass {
    public function identity(): string { return "original"; }
}
class_alias('OriginalClass', 'AliasedClass', true);
$obj = new AliasedClass();
echo $obj->identity(), "\n";
"#
        ),
        vec!["original"]
    );
}

#[test]
fn test_class_alias_instanceof_check() {
    assert_eq!(
        run_prints(
            r#"<?php
class Target {}
class_alias('Target', 'TargetAlias');
$a = new TargetAlias();
echo ($a instanceof Target) ? 'instance_of_target' : 'err', "\n";
"#
        ),
        vec!["instance_of_target"]
    );
}
