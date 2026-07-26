
crate::php_cases! {
    dateperiod_exclude_start_date => {
        r#"<?php
$start = new DateTime('2020-01-01');
$end = new DateTime('2020-01-04');
$interval = new DateInterval('P1D');

$period = new DatePeriod($start, $interval, $end, DatePeriod::EXCLUDE_START_DATE);
$out = [];
foreach ($period as $dt) {
    $out[] = $dt->format('Y-m-d');
}
echo implode(',', $out);
"#,
        ["2020-01-02,2020-01-03"]
    };

    dateperiod_with_recurrences => {
        r#"<?php
$start = new DateTime('2020-01-01');
$interval = new DateInterval('P2D');
$period = new DatePeriod($start, $interval, 2);

$out = [];
foreach ($period as $dt) {
    $out[] = $dt->format('Y-m-d');
}
echo implode(',', $out);
"#,
        ["2020-01-01,2020-01-03,2020-01-05"]
    };
}
