crate::php_cases! {
    curl_escape_and_unescape_strings => {
        r#"<?php
$ch = curl_init();
$escaped = curl_escape($ch, "hello world = +");
$unescaped = curl_unescape($ch, $escaped);

echo $escaped . "|" . $unescaped;
curl_close($ch);
"#,
        ["hello%20world%20%3D%20%2B|hello world = +"]
    };
}
