crate::php_cases! {
    curl_multi_init_and_add_handle => {
        r#"<?php
$mh = curl_multi_init();
$ch1 = curl_init("http://example.com");
$ch2 = curl_init("http://example.org");

$code1 = curl_multi_add_handle($mh, $ch1);
$code2 = curl_multi_add_handle($mh, $ch2);

// 0 is CURLM_OK
echo ($code1 === 0 && $code2 === 0) ? "added" : "failed";

curl_multi_remove_handle($mh, $ch1);
curl_multi_remove_handle($mh, $ch2);
curl_multi_close($mh);
"#,
        ["added"]
    };
}
