
crate::php_cases! {
    str_getcsv_basic => {
        r#"<?php
$str = "apple,banana,orange";
$arr = str_getcsv($str);
echo implode('|', $arr);
"#,
        ["apple|banana|orange"]
    };

    str_getcsv_custom_delimiter_enclosure => {
        r#"<?php
$str = "123;'hello;world';456";
$arr = str_getcsv($str, ';', "'");
echo count($arr) . "|" . $arr[1];
"#,
        ["3|hello;world"]
    };

    str_getcsv_escape_character => {
        r#"<?php
$str = '"field1","field\"2","field3"';
$arr = str_getcsv($str, ',', '"', '\\');
echo implode('|', $arr);
"#,
        ["field1|field\"2|field3"]
    };
}
