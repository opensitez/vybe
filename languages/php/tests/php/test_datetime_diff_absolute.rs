
crate::php_cases! {
    datetime_diff_absolute_true => {
        r#"<?php
$dt1 = new DateTime('2020-01-05');
$dt2 = new DateTime('2020-01-01');

$diff = $dt1->diff($dt2, true);
echo $diff->invert . "|" . $diff->days;
"#,
        ["0|4"]
    };

    datetime_diff_absolute_false => {
        r#"<?php
$dt1 = new DateTime('2020-01-05');
$dt2 = new DateTime('2020-01-01');

$diff = $dt1->diff($dt2, false);
echo $diff->invert . "|" . $diff->days;
"#,
        ["1|4"]
    };

    datetime_diff_absolute_same_day_tz_aware => {
        r#"<?php
$utc = new DateTime('2024-01-01 00:00:00', new DateTimeZone('UTC'));
$ny = new DateTime('2024-01-01 00:00:00', new DateTimeZone('America/New_York'));
$diff = $utc->diff($ny, true);
echo $diff->invert . "|" . $diff->days;
"#,
        ["0|0"]
    };

    datetime_diff_absolute_year_month_parts => {
        r#"<?php
$dt1 = new DateTime('2023-01-15');
$dt2 = new DateTime('2024-03-20');
$diff = $dt1->diff($dt2, true);
echo $diff->y . "," . $diff->m . "," . $diff->d;
"#,
        ["1,2,5"]
    };

    datetime_diff_absolute_time_part => {
        r#"<?php
$dt1 = new DateTime('2024-01-01 10:00:00');
$dt2 = new DateTime('2024-01-01 14:30:10');
$diff = $dt1->diff($dt2, true);
echo $diff->h . "," . $diff->i . "," . $diff->s;
"#,
        ["4,30,10"]
    };

    datetime_diff_absolute_equals_now => {
        r#"<?php
$dt1 = new DateTime('2024-06-15');
$dt2 = new DateTime('2024-06-15');
$diff = $dt1->diff($dt2, true);
echo $diff->invert . "|" . $diff->days;
"#,
        ["0|0"]
    };
}
