
crate::php_cases! {
    mb_decode_numericentity_basic => {
        r#"<?php
$str = "&#8212;";
$convmap = [0x0, 0xffff, 0, 0xffff];
echo mb_decode_numericentity($str, $convmap, "UTF-8") === "—" ? "match" : "fail";
"#,
        ["match"]
    };
}
