
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

    datetime_modify_relative_next_monday => {
        r#"<?php
$dt = new DateTime('2024-01-01');
$dt->modify('next monday');
echo $dt->format('Y-m-d');
"#,
        ["2024-01-08"]
    };

    datetime_modify_relative_first_day_next_month => {
        r#"<?php
$dt = new DateTime('2024-01-15');
$dt->modify('first day of next month');
echo $dt->format('Y-m-d');
"#,
        ["2024-02-01"]
    };

    datetime_modify_relative_last_day_last_month => {
        r#"<?php
$dt = new DateTime('2024-03-10');
$dt->modify('last day of last month');
echo $dt->format('Y-m-d');
"#,
        ["2024-02-29"]
    };

    datetime_modify_relative_midyear_keyword => {
        r#"<?php
$dt = new DateTime('2024-01-10');
$dt->modify('middle of this year');
echo $dt->format('m-d');
"#,
        ["07-02"]
    };

    datetime_modify_relative_invalid => {
        r#"<?php
$dt = new DateTime('2024-01-01');
try {
    $dt->modify('totally-invalid-token');
    echo 'ok';
} catch (Exception $e) {
    echo 'bad';
}
"#,
        ["bad"]
    };
}
