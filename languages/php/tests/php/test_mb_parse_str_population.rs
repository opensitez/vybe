crate::php_cases! {
    mb_parse_str_basic => {
        r#"<?php
mb_parse_str("first=value&arr[]=foo+bar&arr[]=baz", $output);
echo $output['first'] . "|" . $output['arr'][0] . "|" . $output['arr'][1];
"#,
        ["value|foo bar|baz"]
    };
}
