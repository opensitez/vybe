crate::php_cases! {
    nl2br_xhtml_true => {
        r#"<?php
$str = "foo\nbar";
echo nl2br($str);
"#,
        ["foo<br />\nbar"]
    };

    nl2br_xhtml_false => {
        r#"<?php
$str = "foo\r\nbar";
echo nl2br($str, false);
"#,
        ["foo<br>\r\nbar"]
    };
}
