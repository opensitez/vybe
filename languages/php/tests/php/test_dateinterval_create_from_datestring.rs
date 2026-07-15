use super::helpers::run_prints;

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
}
