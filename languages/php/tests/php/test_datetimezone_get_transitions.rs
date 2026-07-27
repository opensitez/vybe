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

    datetimezone_standard_fields => {
        r#"<?php
$tz = new DateTimeZone('UTC');
echo $tz->getName();
echo '|';
echo $tz->getOffset(new DateTime('2020-01-01 00:00:00', $tz));
"#,
        ["UTC|0"]
    };

    datetimezone_locations => {
        r#"<?php
$tz = new DateTimeZone('America/New_York');
$loc = $tz->getLocation();
echo isset($loc['country_code']) ? 'country' : 'none';
echo '|';
echo isset($loc['longitude']) ? 'lon' : 'nolon';
"#,
        ["country|lon"]
    };

    datetimezone_dst_interval_snapshot => {
        r#"<?php
$tz = new DateTimeZone('America/Los_Angeles');
$start = strtotime('2021-03-01 00:00:00');
$end = strtotime('2021-04-01 00:00:00');
$transitions = $tz->getTransitions($start, $end);
echo count($transitions) >= 1 ? 'has' : 'none';
if (count($transitions) > 0 && isset($transitions[0]['isdst'])) {
    echo '|';
    echo $transitions[0]['isdst'] ? 'dst' : 'std';
}
"#,
        ["has|std"]
    };

    datetime_timezone_invalid_zone_throws => {
        r#"<?php
try {
    new DateTimeZone('No_Such_Zone');
    echo 'no';
} catch (Throwable $e) {
    echo 'err';
}
"#,
        ["err"]
    };

    datetimezone_get_transitions_many => {
        r#"<?php
$tz = new DateTimeZone('America/Sao_Paulo');
$start = strtotime('2023-01-01');
$end = strtotime('2023-12-31');
$transitions = $tz->getTransitions($start, $end);
echo is_array($transitions) ? 'ok' : 'bad';
echo '|';
echo count($transitions) > 1 ? 'many' : 'few';
"#,
        ["ok|many"]
    };

    datetimezone_get_transitions_reverse_range => {
        r#"<?php
$tz = new DateTimeZone('Europe/Paris');
$end = strtotime('2024-01-01');
$start = strtotime('2024-06-01');
$transitions = $tz->getTransitions($start, $end);
echo is_array($transitions) ? 'array' : 'na';
"#,
        ["array"]
    };

    datetimezone_get_transition_fields => {
        r#"<?php
$tz = new DateTimeZone('Asia/Kolkata');
$transitions = $tz->getTransitions(strtotime('2024-01-01'), strtotime('2024-12-31'));
$first = $transitions[0];
echo isset($first['ts']) ? 'ts' : 'not_ts';
echo '|';
echo isset($first['offset']) ? 'offset' : 'not_offset';
"#,
        ["ts|offset"]
    };

    datetimezone_get_transition_unknown_abbr => {
        r#"<?php
$tz = new DateTimeZone('Australia/Sydney');
$transitions = $tz->getTransitions(strtotime('2024-01-01'), strtotime('2024-12-31'));
$all = true;
foreach ($transitions as $t) {
    if (!isset($t['abbr']) || !is_string($t['abbr'])) { $all = false; break; }
}
echo $all ? 'abbrs' : 'noabbrs';
"#,
        ["abbrs"]
    };
}
