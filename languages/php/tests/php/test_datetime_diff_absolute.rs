use super::helpers::run_prints;

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
}
