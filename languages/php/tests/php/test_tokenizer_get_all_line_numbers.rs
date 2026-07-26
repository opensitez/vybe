
crate::php_cases! {
    token_get_all_tracks_line_numbers => {
        r#"<?php
$source = "<?php\n\n\$x = 1;\n// comment\n";
$tokens = token_get_all($source);
$lines = [];
foreach ($tokens as $token) {
    if (is_array($token)) {
        $name = token_name($token[0]);
        if ($name === 'T_VARIABLE' || $name === 'T_COMMENT') {
            $lines[] = $name . ':' . $token[2];
        }
    }
}
echo implode(',', $lines);
"#,
        ["T_VARIABLE:3,T_COMMENT:4"]
    };
}
