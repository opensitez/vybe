crate::php_cases! {
    parse_str_basic => {
        r#"<?php
$str = "first=value&arr[]=foo+bar&arr[]=baz";
parse_str($str, $output);
echo $output['first'] . "|" . $output['arr'][0] . "|" . $output['arr'][1];
"#,
        ["value|foo bar|baz"]
    };

    parse_str_nested_arrays => {
        r#"<?php
$str = "user[name]=admin&user[roles][]=editor&user[roles][]=viewer";
parse_str($str, $output);
echo $output['user']['name'] . "|" . implode(',', $output['user']['roles']);
"#,
        ["admin|editor,viewer"]
    };
}
