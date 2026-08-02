<?php
// vybe-test: php/class_inspection/get_class_no_arg_from_method
// origin: languages/php/tests/php/test_class_inspection.rs
// vybe-test-mode: compile

class SelfNaming {
    public function className(): string {
        return get_class($this);
    }
}
$obj = new SelfNaming();
    echo $obj->className();
