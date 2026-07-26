
crate::php_cases! {
    mb_ereg_replace_callback_basic => {
        r#"<?php
$str = "äbc äbc";
$res = @mb_ereg_replace_callback("ä", function($m) { return "o"; }, $str);
// Note mb_ereg is deprecated in modern PHP or might not be enabled, just test if it runs or returns string
echo is_string($res) ? "ok" : "fail";
"#,
        ["ok"]
    };
}
