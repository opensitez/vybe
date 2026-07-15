use super::helpers::run_prints;

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
}
