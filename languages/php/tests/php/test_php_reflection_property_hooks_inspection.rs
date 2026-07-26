use super::helpers::run_prints;

#[test]
fn test_reflection_property_has_hooks() {
    assert_eq!(
        run_prints(
            r#"<?php
class Example {
    public string $name;
}
$rp = new ReflectionProperty(Example::class, 'name');
if (method_exists($rp, 'hasHooks')) {
    echo $rp->hasHooks() ? 'has_hooks' : 'no_hooks', "\n";
} else {
    echo "no_hooks\n";
}
"#
        ),
        vec!["no_hooks"]
    );
}

#[test]
fn test_reflection_property_is_promoted() {
    assert_eq!(
        run_prints(
            r#"<?php
class PromotedDemo {
    public function __construct(public readonly int $id) {}
}
$rp = new ReflectionProperty(PromotedDemo::class, 'id');
echo $rp->isPromoted() ? 'promoted' : 'regular', "\n";
"#
        ),
        vec!["promoted"]
    );
}

#[test]
fn test_reflection_property_is_readonly() {
    assert_eq!(
        run_prints(
            r#"<?php
class ReadonlyDemo {
    public readonly string $uuid;
}
$rp = new ReflectionProperty(ReadonlyDemo::class, 'uuid');
echo $rp->isReadOnly() ? 'readonly' : 'mutable', "\n";
"#
        ),
        vec!["readonly"]
    );
}
