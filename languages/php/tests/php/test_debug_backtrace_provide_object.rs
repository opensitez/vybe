crate::php_cases! {
    debug_backtrace_provide_object => {
        r#"<?php
class TestClass {
    public function doTrace() {
        $trace = debug_backtrace(DEBUG_BACKTRACE_PROVIDE_OBJECT, 1);
        echo isset($trace[0]['object']) && $trace[0]['object'] instanceof TestClass ? "ok" : "fail";
    }
}
$t = new TestClass();
$t->doTrace();
"#,
        ["ok"]
    };
}
