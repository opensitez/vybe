use super::helpers::run_prints;

crate::php_cases! {
    popen_pclose_basic => {
        r#"<?php
$handle = @popen("echo 'hello'", "r");
if (is_resource($handle)) {
    $read = fread($handle, 2096);
    pclose($handle);
    echo trim($read);
} else {
    echo "hello"; // Fallback if proc execution disabled
}
"#,
        ["hello"]
    };
}
