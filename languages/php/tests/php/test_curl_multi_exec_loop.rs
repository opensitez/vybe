
crate::php_cases! {
    curl_multi_exec_loop_execution => {
        r#"<?php
$mh = curl_multi_init();
$ch = curl_init("http://example.com");
curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
curl_multi_add_handle($mh, $ch);

$active = null;
do {
    $status = curl_multi_exec($mh, $active);
    if ($active) {
        curl_multi_select($mh);
    }
} while ($active && $status == CURLM_OK);

$content = curl_multi_getcontent($ch);
echo is_string($content) && strlen($content) > 0 ? "success" : "failed";

curl_multi_remove_handle($mh, $ch);
curl_multi_close($mh);
"#,
        ["success"]
    };
}
