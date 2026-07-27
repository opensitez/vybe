crate::php_cases! {
    mb_convert_variables_basic => {
        r#"<?php
$var1 = "äöü";
$var2 = ["äöü", "test"];
$enc = mb_convert_variables("UTF-8", "ISO-8859-1", $var1, $var2);
echo is_string($enc) ? "ok" : "fail";
"#,
        ["ok"]
    };
}
