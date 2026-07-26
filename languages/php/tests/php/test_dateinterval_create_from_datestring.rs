
crate::php_cases! {
    dateinterval_create_from_datestring => {
        r#"<?php
$i = DateInterval::createFromDateString('1 day + 2 hours');
echo $i->d . "|" . $i->h;
"#,
        ["1|2"]
    };

    dateinterval_create_from_datestring_complex => {
        r#"<?php
$i = DateInterval::createFromDateString('last day of next month');
echo $i->m;
"#,
        ["1"] // Actually this depends on PHP version, but standardly creates relative interval
    };

    dateinterval_create_from_datestring_weeks => {
        r#"<?php
$i = DateInterval::createFromDateString('3 weeks and 2 days');
echo $i->d;
echo '|';
echo $i->d >= 0 ? 'positive' : 'negative';
"#,
        ["23|positive"]
    };

    dateinterval_create_from_datestring_hours_minutes => {
        r#"<?php
$i = DateInterval::createFromDateString('2 hours 30 minutes');
echo $i->h . '|' . $i->i;
"#,
        ["2|30"]
    };

    dateinterval_create_from_datestring_months => {
        r#"<?php
$i = DateInterval::createFromDateString('4 months');
echo $i->m;
"#,
        ["4"]
    };

    dateinterval_create_from_datestring_invalid => {
        r#"<?php
$i = DateInterval::createFromDateString('not a valid duration token');
echo $i === false ? 'false' : 'notfalse';
"#,
        ["notfalse"]
    };
}
