use super::helpers::run_prints;

crate::php_cases! {
    datetime_modify_relative_keywords => {
        r#"<?php
$dt = new DateTime('2020-01-01');
$dt->modify('+1 month');
echo $dt->format('Y-m-d') . "|";
$dt->modify('last day of this month');
echo $dt->format('Y-m-d');
"#,
        ["2020-02-01|2020-02-29"] // 2020 is a leap year
    };
}
