use super::helpers::run_prints;

crate::php_cases! {
    curl_multi_info_read_messages => {
        r#"<?php
$mh = curl_multi_init();
$ch = curl_init("http://example.com");
curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
curl_multi_add_handle($mh, $ch);

$active = null;
do {
    curl_multi_exec($mh, $active);
} while ($active);

$info = curl_multi_info_read($mh);
echo is_array($info) && isset($info['msg']) && $info['msg'] === CURLMSG_DONE ? "done" : "failed";

curl_multi_remove_handle($mh, $ch);
curl_multi_close($mh);
"#,
        ["done"]
    };
}
