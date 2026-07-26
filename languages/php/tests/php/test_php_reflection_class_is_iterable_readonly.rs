use super::helpers::run_prints;

#[test]
fn test_reflection_class_is_iterable() {
    assert_eq!(
        run_prints(
            r#"<?php
class IterableClass implements IteratorAggregate {
    public function getIterator(): Traversable { return new ArrayIterator([]); }
}
$rc = new ReflectionClass(IterableClass::class);
echo $rc->isIterable() ? 'is_iterable' : 'not_iterable', "\n";
"#
        ),
        vec!["is_iterable"]
    );
}

#[test]
fn test_reflection_class_is_cloneable() {
    assert_eq!(
        run_prints(
            r#"<?php
class Unclonable {
    private function __clone() {}
}
$rc = new ReflectionClass(Unclonable::class);
echo $rc->isCloneable() ? 'cloneable' : 'not_cloneable', "\n";
"#
        ),
        vec!["not_cloneable"]
    );
}
