use super::helpers::run_prints;

#[test]
fn test_reflection_class_constant_visibility_modifiers() {
    assert_eq!(
        run_prints(
            r#"<?php
class VisibilityDemo {
    public const PUB = 1;
    protected const PROT = 2;
    private const PRIV = 3;
}
$rc = new ReflectionClass(VisibilityDemo::class);
$pub = $rc->getReflectionConstant('PUB');
$prot = $rc->getReflectionConstant('PROT');
$priv = $rc->getReflectionConstant('PRIV');

echo ($pub->isPublic() ? '1' : '0') . ',' . ($prot->isProtected() ? '1' : '0') . ',' . ($priv->isPrivate() ? '1' : '0'), "\n";
"#
        ),
        vec!["1,1,1"]
    );
}

#[test]
fn test_reflection_class_constant_is_final() {
    assert_eq!(
        run_prints(
            r#"<?php
class FinalConstDemo {
    final public const LOCKED = 'immutable';
}
$rc = new ReflectionClass(FinalConstDemo::class);
$c = $rc->getReflectionConstant('LOCKED');
echo $c->isFinal() ? 'final_const' : 'overrideable', "\n";
"#
        ),
        vec!["final_const"]
    );
}
