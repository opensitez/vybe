crate::php_cases! {
    filter_input_array_basic => {
        r#"<?php
$_GET = ['name' => 'John', 'age' => '30'];
$args = [
    'name' => FILTER_SANITIZE_STRING,
    'age'  => FILTER_VALIDATE_INT,
];
$res = filter_var_array($_GET, $args);
echo $res['name'] . "|" . $res['age'];
"#,
        ["John|30"]
    };
}
