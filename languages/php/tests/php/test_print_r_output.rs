//! `print_r` echo-mode and return-to-string behavior (PHP spec).

crate::php_cases! {
    print_r_scalar_concatenates_with_following_echo => {
        r#"<?php
print_r('scalar');
echo '|';
"#,
        ["scalar|"]
    };

    print_r_integer_concatenates_with_echo => {
        r#"<?php
print_r(99);
echo 'end';
"#,
        ["99end"]
    };

    print_r_return_true_scalar_no_stdout => {
        r#"<?php
$s = print_r('tok', true);
echo $s === 'tok' ? 'match' : 'diff';
"#,
        ["match"]
    };

    print_r_return_true_array_string => {
        r#"<?php
$s = print_r(['k' => 1], true);
echo is_string($s) && str_contains($s, 'k') ? 'ok' : 'fail';
"#,
        ["ok"]
    };

    print_r_return_true_does_not_write_stdout => {
        r#"<?php
$s = print_r('hidden', true);
echo 'marker:' . $s;
"#,
        ["marker:hidden"]
    };

    print_r_array_echo_contains_key_name => {
        r#"<?php
print_r(['id' => 1]);
echo 'END';
"#,
        ["Array\n(\n    [id] => 1\n)END"]
    };
}
