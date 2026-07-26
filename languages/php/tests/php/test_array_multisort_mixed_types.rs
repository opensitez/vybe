
crate::php_cases! {
    array_multisort_mixed_types => {
        r#"<?php
$ar1 = ["10", 11, 100, 100, "a"];
$ar2 = [1, 2, "2", 3, 1];
array_multisort($ar1, SORT_ASC, SORT_STRING,
                $ar2, SORT_NUMERIC, SORT_DESC);
echo implode(',', $ar1) . '|' . implode(',', $ar2);
"#,
        ["10,100,100,11,a|1,3,2,2,1"]
    };

    array_multisort_case_insensitive => {
        r#"<?php
$arr = ["Alpha", "atomic", "Beta", "bank"];
array_multisort($arr, SORT_ASC, SORT_FLAG_CASE | SORT_STRING);
echo implode(',', $arr);
"#,
        ["Alpha,atomic,bank,Beta"]
    };

    array_multisort_preserves_associated_key_correspondence => {
        r#"<?php
$scores = [10, 5, 20, 5];
$names = ["Ann", "Ben", "Cal", "Die"];
array_multisort($scores, SORT_ASC, SORT_NUMERIC, $names, SORT_ASC, SORT_STRING);
echo implode(',', $scores) . "|" . $names[0] . ":" . $names[1] . ":" . $names[2] . ":" . $names[3];
"#,
        ["5,5,10,20|Ben,Die,Ann,Cal"]
    };

    array_multisort_nat_strings_with_regular_numeric_compare => {
        r#"<?php
$versions = ["v2", "v10", "v1", "v3"];
$nums = [1, 2, 3, 4];
array_multisort($versions, SORT_NATURAL | SORT_FLAG_CASE, SORT_NUMERIC, $nums, SORT_DESC, SORT_NUMERIC);
echo implode(',', $versions) . "|" . implode(',', $nums);
"#,
        ["v1,v2,v3,v10|3,1,4,2"]
    };

    array_multisort_with_duplicates_and_reverse => {
        r#"<?php
$vals = ["b", "a", "b", "a", "c"];
$weights = [3, 1, 2, 4, 0];
array_multisort($vals, SORT_DESC, SORT_STRING, $weights, SORT_ASC, SORT_NUMERIC);
echo implode(',', $vals) . "|" . implode(',', $weights);
"#,
        ["c,b,b,a,a|0,2,3,1,4"]
    };

    array_multisort_regular_and_nat_combined => {
        r#"<?php
$a = ["10", "2", "1", "3"];
$b = ["y", "x", "z", "w"];
array_multisort($a, SORT_ASC, SORT_STRING, $b, SORT_DESC, SORT_NUMERIC);
echo implode(',', $a) . "|" . implode(',', $b);
"#,
        ["1,10,2,3|z,y,x,w"]
    };

    array_multisort_mixed_case_strings => {
        r#"<?php
$keys = ["beta", "Alpha", "charlie", "Bravo"];
$names = ["B", "A", "C", "D"];
array_multisort($keys, SORT_ASC, SORT_STRING | SORT_FLAG_CASE, $names, SORT_ASC, SORT_STRING);
echo implode(',', $keys) . "|" . implode(',', $names);
"#,
        ["Alpha,Bravo,beta,charlie|A,D,B,C"]
    };

    array_multisort_by_floats_and_strings => {
        r#"<?php
$vals = ["1.2", 3, 2.5, "2.5", 1];
$labels = ["a", "b", "c", "d", "e"];
array_multisort($vals, SORT_ASC, SORT_NUMERIC, $labels, SORT_ASC, SORT_STRING);
echo implode(',', $vals) . "|" . implode(',', $labels);
"#,
        ["1,2.5,2.5,3,1.2|e,c,d,b,a"]
    };

    array_multisort_assoc_tie_breaker_preserves_payload_by_index => {
        r#"<?php
$scores = [10, 10, 10, 20];
$ids = [5, 2, 9, 1];
$names = ["x", "y", "z", "w"];
array_multisort($scores, SORT_ASC, SORT_NUMERIC, $ids, SORT_ASC, SORT_NUMERIC, $names, SORT_ASC, SORT_STRING);
echo implode(',', $scores) . "|" . implode(',', $ids) . "|" . implode(',', $names);
"#,
        ["10,10,10,20|2,5,9,1|y,x,z,w"]
    };

    array_multisort_large_and_edge_flags => {
        r#"<?php
$flags = [SORT_ASC, SORT_DESC];
$primary = ["a", "b", "c", "d"];
$secondary = [4, 3, 2, 1];
array_multisort($primary, $flags[0], SORT_STRING, $secondary, $flags[1], SORT_NUMERIC);
echo implode(',', $primary) . "|" . implode(',', $secondary);
"#,
        ["a,b,c,d|4,3,2,1"]
    };

    array_multisort_nested_array_inputs => {
        r#"<?php
$matrixA = [
    [1, "3"],
    [2, "1"],
    [0, "2"]
];
$matrixB = ["x", "y", "z"];
array_multisort(array_column($matrixA, 1), SORT_NUMERIC, SORT_ASC, array_column($matrixA, 0), SORT_ASC, SORT_NUMERIC, $matrixB, SORT_ASC, SORT_STRING);
echo implode(',', $matrixB);
"#,
        ["y,z,x"]
    };

    array_multisort_with_empty_second_array_still_reorders_first => {
        r#"<?php
$scores = [8, 2, 5];
$payload = [];
array_multisort($scores, SORT_DESC, SORT_NUMERIC, $payload);
echo implode(',', $scores);
"#,
        ["8,5,2"]
    };
}
