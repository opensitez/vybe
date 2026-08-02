<?php
// vybe-test: php/php_var_dump_export_debug_info/test_php_debug_backtrace_provide_object_option
// origin: languages/php/tests/php/test_php_var_dump_export_debug_info.rs
// vybe-test-mode: compile

class Inspector {
    public function inspect() {
        return debug_backtrace(DEBUG_BACKTRACE_PROVIDE_OBJECT);
    }
}

$i = new Inspector();
$trace = $i->inspect();
echo isset($trace[0]["object"]) && $trace[0]["object"] === $i ? "OBJECT_PROVIDED" : "FAIL";
