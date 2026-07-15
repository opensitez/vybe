use super::helpers::run_prints;

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
}
