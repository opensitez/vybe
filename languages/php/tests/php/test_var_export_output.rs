//! `var_export` echo-mode and return-to-string behavior (PHP spec).

crate::php_cases! {
    var_export_echo_mode_scalar_concatenates => {
        r#"<?php
var_export(42);
echo '|';
"#,
        ["42|"]
    };

    var_export_echo_mode_bool_false => {
        r#"<?php
var_export(false);
echo '|';
"#,
        ["false|"]
    };

    var_export_return_true_assoc_array => {
        r#"<?php
$s = var_export(['a' => 1], true);
echo str_contains($s, "'a'") ? 'has-key' : 'missing';
"#,
        ["has-key"]
    };

    var_export_return_true_nested_array => {
        r#"<?php
$s = var_export(['outer' => ['inner' => 2]], true);
echo str_contains($s, 'inner') ? 'nested' : 'flat';
"#,
        ["nested"]
    };
}
