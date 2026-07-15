use super::helpers::run_prints;

crate::php_cases! {
    array_reduce_with_initial_value => {
        r#"<?php
$a = [1, 2, 3, 4, 5];
$res = array_reduce($a, function($carry, $item) {
    $carry *= $item;
    return $carry;
}, 10);
echo $res;
"#,
        ["1200"]
    };

    array_reduce_empty_array_with_initial => {
        r#"<?php
$res = array_reduce([], function($c, $i) { return $c + $i; }, "initial");
echo $res;
"#,
        ["initial"]
    };

    array_reduce_empty_array_no_initial => {
        r#"<?php
$res = array_reduce([], function($c, $i) { return $c + $i; });
echo is_null($res) ? "null" : "not null";
"#,
        ["null"]
    };
}
