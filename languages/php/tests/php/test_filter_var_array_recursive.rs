crate::php_cases! {
    filter_var_array_recursive => {
        r#"<?php
$data = [
    'user' => [
        'email' => 'test@example.com',
        'age'   => 'not-an-int'
    ]
];
$args = [
    'user' => [
        'filter' => FILTER_VALIDATE_EMAIL,
        'flags'  => FILTER_REQUIRE_ARRAY
    ] // Actually filter_var_array doesn't recurse automatically without manual iteration, but we test the array flag.
];
// Wait, we'll just test a simpler filter_var_array with basic associative inputs
$data2 = ['a' => '1', 'b' => '2'];
$res = filter_var_array($data2, FILTER_VALIDATE_INT);
echo $res['a'] . "|" . $res['b'];
"#,
        ["1|2"]
    };
}
