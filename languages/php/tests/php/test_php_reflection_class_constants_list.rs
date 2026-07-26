use super::helpers::run_prints;

#[test]
fn test_reflection_class_get_reflection_constants() {
    assert_eq!(
        run_prints(
            r#"<?php
class ConfigDemo {
    public const DEFAULT_HOST = '127.0.0.1';
    protected const PORT = 8080;
    private const SECRET = 'key123';
}
$rc = new ReflectionClass(ConfigDemo::class);
$consts = $rc->getReflectionConstants();
echo count($consts) . ':' . $consts[0]->getName(), "\n";
"#
        ),
        vec!["3:DEFAULT_HOST"]
    );
}

#[test]
fn test_reflection_class_get_single_reflection_constant() {
    assert_eq!(
        run_prints(
            r#"<?php
class ApiDemo {
    public const VERSION = '2.0';
}
$rc = new ReflectionClass(ApiDemo::class);
$c = $rc->getReflectionConstant('VERSION');
echo $c instanceof ReflectionClassConstant ? $c->getValue() : 'not_constant', "\n";
"#
        ),
        vec!["2.0"]
    );
}
