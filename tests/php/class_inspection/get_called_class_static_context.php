<?php
// vybe-test: php/class_inspection/get_called_class_static_context
// origin: languages/php/tests/php/test_class_inspection.rs
// vybe-test-mode: compile

class ParentClass {
    public static function whoAmI(): string {
        return get_called_class();
    }
}
class ChildClass extends ParentClass {}
echo ChildClass::whoAmI();
