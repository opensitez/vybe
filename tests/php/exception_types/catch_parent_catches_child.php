<?php
// vybe-test: php/exception_types/catch_parent_catches_child
// origin: languages/php/tests/php/test_exception_types.rs
// vybe-test-mode: compile

class BaseException extends Exception {}
class ChildException extends BaseException {}
try {
    throw new ChildException('child');
} catch (BaseException $e) {
    echo $e->getMessage();
}
