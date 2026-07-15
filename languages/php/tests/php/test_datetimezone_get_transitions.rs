use super::helpers::run_prints;

crate::php_cases! {
    datetimezone_get_transitions => {
        r#"<?php
$tz = new DateTimeZone('Europe/London');
$transitions = $tz->getTransitions(
    strtotime('2020-03-25'),
    strtotime('2020-04-05')
);

// We expect a DST transition on 2020-03-29
echo count($transitions) . "|";
if (count($transitions) > 1) {
    echo $transitions[1]['isdst'] ? 'DST' : 'STD';
}
"#,
        ["2|DST"]
    };
}
