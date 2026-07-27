crate::php_cases! {
    mb_stristr_basic => {
        r#"<?php
$str = "ÄÖÜ test";
echo mb_stristr($str, "öü", true, "UTF-8");
"#,
        ["Ä"] // Returns before needle
    };
}
