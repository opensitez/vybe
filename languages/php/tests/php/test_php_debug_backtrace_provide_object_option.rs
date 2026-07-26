use super::helpers::run_prints;

#[test]
fn test_debug_backtrace_provide_object() {
    assert_eq!(
        run_prints(
            r#"<?php
class TraceDemo {
    public string $id = 'demo_obj';
    public function traceSelf() {
        return debug_backtrace(DEBUG_BACKTRACE_PROVIDE_OBJECT, 1)[0];
    }
}
$td = new TraceDemo();
$frame = $td->traceSelf();
echo (isset($frame['object']) && $frame['object'] === $td) ? 'object_provided' : 'err', "\n";
"#
        ),
        vec!["object_provided"]
    );
}
