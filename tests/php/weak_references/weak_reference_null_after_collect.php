<?php
// vybe-test: php/weak_references/weak_reference_null_after_collect
// origin: languages/php/tests/php/test_weak_references.rs
// vybe-test-mode: compile

class Temp {}
$weak = null;
{
    $obj = new Temp();
    $weak = WeakReference::create($obj);
    unset($obj);
}
$result = $weak->get();
echo $result === null ? 'null' : 'alive';
