
crate::php_cases! {
    curl_share_init_and_setopt => {
        r#"<?php
$sh = curl_share_init();
$res1 = curl_share_setopt($sh, CURLSHOPT_SHARE, CURL_LOCK_DATA_COOKIE);
$res2 = curl_share_setopt($sh, CURLSHOPT_SHARE, CURL_LOCK_DATA_DNS);

echo ($res1 && $res2) ? "shared" : "failed";

$ch = curl_init("http://example.com");
curl_setopt($ch, CURLOPT_SHARE, $sh);
curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
curl_exec($ch);

curl_share_close($sh);
"#,
        ["shared"]
    };
}
