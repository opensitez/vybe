crate::php_cases! {
    str_word_count_return_formats => {
        r#"<?php
$str = "Hello fri3nd, you're looking good!";
echo str_word_count($str) . "|";
echo count(str_word_count($str, 1)) . "|";
echo implode(',', array_keys(str_word_count($str, 2)));
"#,
        ["5|5|0,6,14,21,29"]
    };

    str_word_count_char_lists => {
        r#"<?php
$str = "Hello fri3nd, you're looking good!";
echo count(str_word_count($str, 1, '3'));
"#,
        ["6"]
    };
}
