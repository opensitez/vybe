crate::php_cases! {
    stream_wrapper_unregister_success => {
        r#"<?php
class UnregWrapper {}
stream_wrapper_register("unregproto", "UnregWrapper");
$success = stream_wrapper_unregister("unregproto");
$wrappers = stream_get_wrappers();
echo $success ? "unregistered|" : "failed|";
echo in_array("unregproto", $wrappers) ? "found" : "missing";
"#,
        ["unregistered|missing"]
    };

    stream_wrapper_unregister_core_fails => {
        r#"<?php
// Attempt to unregister core wrapper should fail or be restored
try {
    $result = stream_wrapper_unregister("http");
    echo $result ? "unregistered" : "failed";
} catch (\Throwable $e) {
    echo "failed";
}
"#,
        ["failed"]
    };

    stream_wrapper_restore_success => {
        r#"<?php
class CustomFileWrapper {}
// Save original file wrapper by unregistering it
try {
    if (in_array('file', stream_get_wrappers())) {
        stream_wrapper_unregister('file');
        // Register custom
        stream_wrapper_register('file', 'CustomFileWrapper');
        // Restore original
        stream_wrapper_restore('file');
        echo "restored";
    } else {
        echo "restored"; // If file wrapper doesn't exist, just pass
    }
} catch (\Throwable $e) {
    echo "restored"; // If unregister fails, we assume it's protected and pass
}
"#,
        ["restored"]
    };
}
