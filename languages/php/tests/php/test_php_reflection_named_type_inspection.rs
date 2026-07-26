use super::helpers::run_prints;

#[test]
fn test_reflection_named_type_get_name() {
    assert_eq!(
        run_prints(
            r#"<?php
function process(int $x): void {}
$rf = new ReflectionFunction('process');
$param = $rf->getParameters()[0];
$type = $param->getType();
echo $type instanceof ReflectionNamedType ? $type->getName() : 'not_named', "\n";
"#
        ),
        vec!["int"]
    );
}

#[test]
fn test_reflection_named_type_is_builtin() {
    assert_eq!(
        run_prints(
            r#"<?php
class CustomObj {}
function handle(CustomObj $obj, string $str): void {}
$rf = new ReflectionFunction('handle');
$p1 = $rf->getParameters()[0]->getType();
$p2 = $rf->getParameters()[1]->getType();
echo ($p1->isBuiltin() ? 'p1_builtin' : 'p1_class') . ',' . ($p2->isBuiltin() ? 'p2_builtin' : 'p2_class'), "\n";
"#
        ),
        vec!["p1_class,p2_builtin"]
    );
}
