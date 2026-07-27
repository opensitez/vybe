crate::php_cases! {
    array_walk_recursive_by_reference => {
        r#"<?php
$sweet = ['a' => 'apple', 'b' => 'banana'];
$fruits = ['sweet' => $sweet, 'sour' => 'lemon'];

function test_print(&$item, $key, $prefix) {
    $item = "$prefix: $item";
}

array_walk_recursive($fruits, 'test_print', 'fruit');

echo $fruits['sweet']['a'] . "|" . $fruits['sour'];
"#,
        ["fruit: apple|fruit: lemon"]
    };

    array_walk_recursive_objects => {
        r#"<?php
class Obj { public $val = 1; }
$arr = [new Obj(), [new Obj()]];

array_walk_recursive($arr, function($v, $k) {
    $v->val += 10;
});
echo $arr[0]->val . "|" . $arr[1][0]->val;
"#,
        ["11|11"]
    };

    array_walk_recursive_numeric_keys => {
        r#"<?php
$a = ['x' => ['n' => 1], 'y' => ['n' => 2], 'z' => [3,4]];
$keys = [];
array_walk_recursive($a, function($v, $k) use (&$keys) {
    if (is_int($v)) {
        $keys[] = $k;
    }
});
echo implode(',', $keys);
"#,
        ["n,n,0,1"]
    };

    array_walk_recursive_mutates_strings => {
        r#"<?php
$a = ['a' => ['value' => 'x'], 'b' => ['value' => 'y']];
array_walk_recursive($a, function(&$v) {
    if (is_string($v)) {
        $v .= '!';
    }
});
echo $a['a']['value'] . '|' . $a['b']['value'];
"#,
        ["x!|y!"]
    };

    array_walk_recursive_with_user_data => {
        r#"<?php
$a = [1, 2];
$sum = 0;
array_walk_recursive($a, function($v, $k, $factor) use (&$sum) {
    $sum += $v * $factor;
}, 5);
echo $sum;
"#,
        ["15"]
    };

    array_walk_recursive_empty_array => {
        r#"<?php
$a = [];
$seen = 0;
$result = array_walk_recursive($a, function($value, $key) use (&$seen) {
    $seen++;
});
echo ($result ? '1' : '0') . '|' . $seen;
"#,
        ["1|0"]
    };

    array_walk_recursive_top_level_scalars => {
        r#"<?php
$a = ['a' => 1, 'b' => 2, 'c' => 'three'];
$sum = 0;
array_walk_recursive($a, function($v, $k) use (&$sum) {
    $sum += is_numeric($v) ? (int)$v : 0;
});
echo $sum;
"#,
        ["3"]
    };

    array_walk_recursive_nested_reference_mutation => {
        r#"<?php
$a = ['left' => ['value' => 2], 'right' => ['value' => 3]];
array_walk_recursive($a, function(&$v, $k) {
    if (is_int($v)) {
        $v += 1;
    }
});
echo $a['left']['value'] . '|' . $a['right']['value'];
"#,
        ["3|4"]
    };

    array_walk_recursive_user_data_and_keys => {
        r#"<?php
$a = ['outer' => ['x' => 5], 'inner' => ['y' => 6]];
$out = [];
array_walk_recursive($a, function($v, $k, $prefix) use (&$out) {
    if (is_int($v)) {
        $out[] = $prefix . ':' . $k;
    }
}, 'K');
echo implode('|', $out);
"#,
        ["K:x|K:y"]
    };

    array_walk_recursive_nested_scalar_collection => {
        r#"<?php
$a = ['a' => [1, 2, 3], 'b' => ['c' => 4], 'd' => true];
$leaves = [];
array_walk_recursive($a, function($v, $k) use (&$leaves) {
    if (is_scalar($v)) {
        $leaves[] = $k . '=' . (string)$v;
    }
});
echo implode('|', $leaves);
"#,
        ["0=1|1=2|2=3|c=4|d=1"]
    };

    array_walk_recursive_string_keys_with_numeric_text => {
        r#"<?php
$a = ['0' => 'one', 1 => 'two', '01' => 'three'];
$keys = [];
array_walk_recursive($a, function($v, $k) use (&$keys) {
    if (is_string($k) || is_int($k)) {
        $keys[] = $k;
    }
});
echo implode(',', $keys);
"#,
        ["0,1,01"]
    };

    array_walk_recursive_exception_propagates => {
        r#"<?php
$a = ['ok' => 1, 'bad' => 2];
try {
    array_walk_recursive($a, function($v) {
        if ($v === 2) {
            throw new Exception('boom');
        }
    });
    echo 'no-exception';
} catch (Exception $e) {
    echo $e->getMessage();
}
"#,
        ["boom"]
    };

    array_walk_recursive_can_set_null => {
        r#"<?php
$a = ['x' => ['v' => 10], 'y' => ['v' => 20], 'z' => 'keep'];
array_walk_recursive($a, function(&$v) {
    if (is_int($v)) {
        $v = null;
    }
});
echo (($a['x']['v'] === null ? '1' : '0') . ($a['y']['v'] === null ? '1' : '0') . ($a['z'] === null ? '1' : '0'));
"#,
        ["110"]
    };
}
