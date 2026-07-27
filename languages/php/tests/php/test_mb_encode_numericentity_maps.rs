crate::php_cases! {
    mb_encode_numericentity_basic => {
        r#"<?php
$str = "—"; // EM DASH U+2014 (8212)
$convmap = [0x0, 0xffff, 0, 0xffff];
echo mb_encode_numericentity($str, $convmap, "UTF-8");
"#,
        ["&#8212;"]
    };
}
