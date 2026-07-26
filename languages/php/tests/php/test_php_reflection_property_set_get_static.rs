use super::helpers::run_prints;

#[test]
fn test_reflection_property_set_get_value_instance() {
    assert_eq!(
        run_prints(
            r#"<?php
class Container {
    public int $value = 10;
}
$c = new Container();
$rp = new ReflectionProperty(Container::class, 'value');
$rp->setValue($c, 50);
echo $rp->getValue($c), "\n";
"#
        ),
        vec!["50"]
    );
}

#[test]
fn test_reflection_property_set_get_value_static() {
    assert_eq!(
        run_prints(
            r#"<?php
class GlobalState {
    public static string $env = 'dev';
}
$rp = new ReflectionProperty(GlobalState::class, 'env');
$rp->setValue(null, 'prod');
echo $rp->getValue(), "\n";
"#
        ),
        vec!["prod"]
    );
}
