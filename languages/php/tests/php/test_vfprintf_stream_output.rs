crate::php_cases! {
    vfprintf_basic => {
        r#"<?php
$fp = fopen("php://memory", "w+");
$format = "Name: %s, Age: %d";
$args = ["Alice", 30];

$len = vfprintf($fp, $format, $args);
rewind($fp);
echo stream_get_contents($fp) . "|" . $len;
fclose($fp);
"#,
        ["Name: Alice, Age: 30|21"]
    };
}
