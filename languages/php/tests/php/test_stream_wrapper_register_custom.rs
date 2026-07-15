use super::helpers::run_prints;

crate::php_cases! {
    stream_wrapper_register_success => {
        r#"<?php
class MyWrapper {
    public function stream_open($path, $mode, $options, &$opened_path) {
        return true;
    }
}
$result = stream_wrapper_register("myproto", "MyWrapper");
echo $result ? "registered" : "failed";
"#,
        ["registered"]
    };

    stream_wrapper_register_duplicate_fails => {
        r#"<?php
class MyWrapper2 {
    public function stream_open($path, $mode, $options, &$opened_path) {
        return true;
    }
}
stream_wrapper_register("myproto2", "MyWrapper2");
// Attempting to register the same protocol again should return false or throw
try {
    $result = stream_wrapper_register("myproto2", "MyWrapper2");
    echo $result ? "registered" : "failed";
} catch (\Exception $e) {
    echo "failed";
} catch (\Error $e) {
    echo "failed";
}
"#,
        ["failed"]
    };

    stream_wrapper_register_invalid_class_fails => {
        r#"<?php
try {
    $result = stream_wrapper_register("myproto3", "NonExistentClass");
    echo $result ? "registered" : "failed";
} catch (\Throwable $e) {
    echo "failed";
}
"#,
        ["failed"]
    };

    stream_get_wrappers_contains_custom => {
        r#"<?php
class MyWrapper4 {}
stream_wrapper_register("myproto4", "MyWrapper4");
$wrappers = stream_get_wrappers();
echo in_array("myproto4", $wrappers) ? "found" : "missing";
"#,
        ["found"]
    };
}
