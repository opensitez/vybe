
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

    array_reduce_without_initial_singleton => {
        r#"<?php
$res = array_reduce([42], fn($carry, $item) => $carry + $item);
echo $res;
"#,
        ["42"]
    };

    array_reduce_callback_default_to_null => {
        r#"<?php
$res = array_reduce([1, 2, 3], function($carry, $item) {
    $item = $item + 1; // side-effect to keep optimizer from folding
});
echo is_null($res) ? 'null' : $res;
"#,
        ["null"]
    };

    array_reduce_with_keyless_string_concat => {
        r#"<?php
$res = array_reduce(['a', 'b', 'c'], fn($carry, $item) => $carry . $item, '');
echo $res;
"#,
        ["abc"]
    };

    array_reduce_building_object_like_array => {
        r#"<?php
$res = array_reduce(
    [1, 2, 3],
    function($carry, $item) {
        $carry[] = $item * 2;
        return $carry;
    },
    []
);
echo implode('|', $res);
"#,
        ["2|4|6"]
    };

    array_reduce_uses_first_element_when_no_initial => {
        r#"<?php
$res = array_reduce([10, 20, 30], function($carry, $item) {
    return $carry + $item;
});
echo $res;
"#,
        ["60"]
    };

    array_reduce_with_boolean_seed => {
        r#"<?php
$res = array_reduce([1, 0, 1], function($carry, $item) {
    return $carry && (bool)$item;
}, true);
echo $res ? 'true' : 'false';
"#,
        ["false"]
    };

    array_reduce_with_assoc_keys => {
        r#"<?php
$res = array_reduce(
    ['x' => 1, 'y' => 2, 'z' => 3],
    function($carry, $item) {
        return ($carry === null ? '' : $carry . ',') . $item;
    },
    null
);
echo $res;
"#,
        ["1,2,3"]
    };

    array_reduce_empty_numeric_seed_zero => {
        r#"<?php
echo array_reduce([], fn($carry, $item) => $carry + $item, 0);
"#,
        ["0"]
    };

    array_reduce_with_non_scalar_seed => {
        r#"<?php
$seed = (object)['count' => 0];
$res = array_reduce([1, 2, 3], function($carry, $item) {
    $carry->count += $item;
    return $carry;
}, $seed);
echo $res->count;
"#,
        ["6"]
    };
}
