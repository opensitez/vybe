crate::php_cases! {
    dateinterval_format_specifiers => {
        r#"<?php
$i = new DateInterval('P2Y4DT6H8M');
echo $i->format('%y years, %d days, %h hours, %i minutes');
"#,
        ["2 years, 4 days, 6 hours, 8 minutes"]
    };

    dateinterval_format_zero_padding => {
        r#"<?php
$i = new DateInterval('P1M3D');
echo $i->format('%M months %D days');
"#,
        ["01 months 03 days"]
    };

    dateinterval_format_inverted => {
        r#"<?php
$start = new DateTime('2024-01-01');
$end = new DateTime('2023-01-01');
$diff = $end->diff($start);
echo $diff->invert . '|' . $diff->format('%R%y years');
"#,
        ["0|+1 years"]
    };

    dateinterval_format_full_pattern => {
        r#"<?php
$i = new DateInterval('P3Y2M1DT4H5M6S');
echo $i->format('%Y years %M months %D days %H hours %I minutes %S seconds');
"#,
        ["03 years 02 months 01 days 04 hours 05 minutes 06 seconds"]
    };

    dateinterval_format_time_only => {
        r#"<?php
$i = new DateInterval('PT90S');
echo $i->format('%i:%s');
"#,
        ["00:90"]
    };

    dateinterval_format_from_iso_days => {
        r#"<?php
$i = new DateInterval('P15D');
echo $i->format('%R%a');
"#,
        ["+15"]
    };
}
