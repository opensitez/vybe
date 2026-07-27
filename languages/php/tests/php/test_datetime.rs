use super::helpers::run_prints;

// ── date() / time() basics ───────────────────────────────────────
#[test]
fn time_returns_integer() {
    assert_eq!(
        run_prints(
            r#"<?php
$t = time();
echo is_int($t) ? "yes" : "no";
"#
        ),
        &["yes"]
    );
}

#[test]
fn date_format_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$ts = mktime(14, 30, 0, 6, 15, 2024);
echo date("Y-m-d", $ts);
echo date("H:i:s", $ts);
"#
        ),
        // echo emits no newline, so PHP concatenates the two into one line.
        &["2024-06-1514:30:00"]
    );
}

#[test]
fn date_format_parts() {
    assert_eq!(
        run_prints(
            r#"<?php
$ts = mktime(0, 0, 0, 12, 25, 2023);
echo date("Y", $ts);
echo date("m", $ts);
echo date("d", $ts);
echo date("D", $ts);
"#
        ),
        // echo emits no newline: all four parts concatenate into one line.
        &["20231225Mon"]
    );
}

#[test]
fn mktime_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$ts1 = mktime(0, 0, 0, 1, 1, 2020);
$ts2 = mktime(0, 0, 0, 1, 2, 2020);
echo $ts2 - $ts1;
"#
        ),
        &["86400"]
    );
}

#[test]
fn date_default_timezone_set_utc_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo date_default_timezone_set('UTC') ? 'yes' : 'no';
"#
        ),
        &["yes"]
    );
}

// ── strtotime ────────────────────────────────────────────────────
#[test]
fn strtotime_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$ts = strtotime("2024-01-15");
echo date("Y-m-d", $ts);
"#
        ),
        &["2024-01-15"]
    );
}

#[test]
fn strtotime_relative() {
    assert_eq!(
        run_prints(
            r#"<?php
$base = strtotime("2024-06-15");
$next = strtotime("+7 days", $base);
echo date("Y-m-d", $next);
$prev = strtotime("-1 month", $base);
echo date("Y-m-d", $prev);
"#
        ),
        // echo emits no newline: both dates concatenate into one line.
        &["2024-06-222024-05-15"]
    );
}

// ── checkdate ────────────────────────────────────────────────────
#[test]
fn checkdate_valid() {
    assert_eq!(
        run_prints(
            r#"<?php
echo checkdate(2, 29, 2024) ? "valid" : "invalid";
echo checkdate(2, 29, 2023) ? "valid" : "invalid";
echo checkdate(13, 1, 2024) ? "valid" : "invalid";
echo checkdate(12, 31, 2024) ? "valid" : "invalid";
"#
        ),
        &["validinvalidinvalidvalid"]
    );
}

// ── DateTime class ───────────────────────────────────────────────
#[test]
fn datetime_construct() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime("2024-06-15 14:30:00");
echo $dt->format("Y-m-d");
echo $dt->format("H:i:s");
"#
        ),
        &["2024-06-1514:30:00"]
    );
}

#[test]
fn datetime_modify() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime("2024-01-01");
$dt->modify("+6 months");
echo $dt->format("Y-m-d");
$dt->modify("-10 days");
echo $dt->format("Y-m-d");
"#
        ),
        &["2024-07-012024-06-21"]
    );
}

#[test]
fn datetime_diff() {
    assert_eq!(
        run_prints(
            r#"<?php
$d1 = new DateTime("2024-01-01");
$d2 = new DateTime("2024-03-01");
$diff = $d1->diff($d2);
echo $diff->days;
echo $diff->m;
"#
        ),
        &["602"]
    );
}

#[test]
fn datetime_timezone_from_constructor_and_name() {
    assert_eq!(
        run_prints(
            r#"<?php
$tz = new DateTimeZone('America/New_York');
$dt = new DateTime('2024-01-02 15:00:00', $tz);
echo $dt->format('e');
echo $dt->getTimezone()->getName();
"#,
        ),
        &["America/New_YorkAmerica/New_York"]
    );
}

#[test]
fn datetime_set_timezone_with_copy() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime('2024-01-02 12:00:00', new DateTimeZone('UTC'));
$other = clone $dt;
$other->setTimezone(new DateTimeZone('Asia/Tokyo'));
echo $dt->format('H:i');
echo '|';
echo $other->format('H:i');
"#,
        ),
        &["12:00|21:00"]
    );
}

#[test]
fn datetime_timezone_offset_diff() {
    assert_eq!(
        run_prints(
            r#"<?php
$utc = new DateTimeImmutable('2024-01-02 12:00:00', new DateTimeZone('UTC'));
$ny = $utc->setTimezone(new DateTimeZone('America/New_York'));
echo $utc->getOffset();
echo '|';
echo $ny->getOffset();
"#,
        ),
        &["0|-18000"]
    );
}

#[test]
fn datetime_interval_month_boundary_rolls() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime('2024-01-31');
$dt->add(new DateInterval('P1M'));
echo $dt->format('Y-m-d');
$dt->sub(new DateInterval('P1M'));
echo $dt->format('Y-m-d');
"#,
        ),
        &["2024-03-02 2024-01-31"]
    );
}

#[test]
fn datetime_diff_with_negative_interval_sign() {
    assert_eq!(
        run_prints(
            r#"<?php
$start = new DateTime('2024-06-10');
$end = new DateTime('2024-06-01');
$diff = $start->diff($end);
echo $diff->invert ? 'neg' : 'pos';
echo '|';
echo $diff->days;
"#,
        ),
        &["neg|9"]
    );
}

#[test]
fn datetime_format_various() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime("2024-12-25 09:30:45");
echo $dt->format("l");
echo $dt->format("F j, Y");
echo $dt->format("g:i A");
"#
        ),
        &["WednesdayDecember 25, 20249:30 AM"]
    );
}

#[test]
fn datetime_create_from_format() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = DateTime::createFromFormat("d/m/Y", "25/12/2024");
echo $dt->format("Y-m-d");
"#
        ),
        &["2024-12-25"]
    );
}

#[test]
fn datetime_get_timestamp() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime("2024-01-01 00:00:00");
$ts = $dt->getTimestamp();
echo is_int($ts) ? "yes" : "no";
echo date("Y", $ts);
"#
        ),
        &["yes2024"]
    );
}

// ── DateInterval ─────────────────────────────────────────────────
#[test]
fn dateinterval_construct() {
    assert_eq!(
        run_prints(
            r#"<?php
$interval = new DateInterval("P1Y2M3D");
echo $interval->y;
echo $interval->m;
echo $interval->d;
"#
        ),
        &["123"]
    );
}

#[test]
fn datetime_add_interval() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime("2024-01-01");
$interval = new DateInterval("P30D");
$dt->add($interval);
echo $dt->format("Y-m-d");
"#
        ),
        &["2024-01-31"]
    );
}

#[test]
fn datetime_sub_interval() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime("2024-06-15");
$interval = new DateInterval("P3M");
$dt->sub($interval);
echo $dt->format("Y-m-d");
"#
        ),
        &["2024-03-15"]
    );
}

// ── getdate / localtime ──────────────────────────────────────────
#[test]
fn getdate_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$ts = mktime(14, 30, 0, 6, 15, 2024);
$info = getdate($ts);
echo $info["year"];
echo $info["mon"];
echo $info["mday"];
echo $info["hours"];
"#
        ),
        &["202461514"]
    );
}

// ── Date arithmetic patterns ─────────────────────────────────────
#[test]
fn days_between_dates() {
    assert_eq!(
        run_prints(
            r#"<?php
$start = new DateTime("2024-01-01");
$end = new DateTime("2024-12-31");
$diff = $start->diff($end);
echo $diff->days;
"#
        ),
        &["365"]
    );
}

#[test]
fn date_comparison() {
    assert_eq!(
        run_prints(
            r#"<?php
$d1 = new DateTime("2024-01-01");
$d2 = new DateTime("2024-06-15");
echo $d1 < $d2 ? "before" : "after";
echo $d1 == $d2 ? "equal" : "not equal";
"#
        ),
        &["beforenot equal"]
    );
}

#[test]
fn date_immutable() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTimeImmutable("2024-01-01");
$modified = $dt->modify("+1 month");
echo $dt->format("Y-m-d");
echo $modified->format("Y-m-d");
"#
        ),
        // echo emits no newline: both dates concatenate into one line.
        &["2024-01-012024-02-01"]
    );
}

#[test]
fn test_checkdate_leap_year() {
    assert_eq!(
        run_prints(
            r#"<?php
echo checkdate(2, 29, 2024) ? 'leap_ok' : 'err';
echo checkdate(2, 29, 2023) ? 'err' : ' non_leap_ok';
"#
        ),
        &["leap_ok non_leap_ok"]
    );
}

#[test]
fn test_gmmktime_gmt_timestamp() {
    assert_eq!(
        run_prints(
            r#"<?php
$gmt = gmmktime(0, 0, 0, 1, 1, 2024);
echo date("Y-m-d", $gmt);
"#
        ),
        &["2024-01-01"]
    );
}

#[test]
fn test_idate_single_character_integer() {
    assert_eq!(
        run_prints(
            r#"<?php
$ts = mktime(15, 45, 0, 10, 25, 2024);
echo idate('Y', $ts) . ':' . idate('m', $ts) . ':' . idate('d', $ts);
"#
        ),
        &["2024:10:25"]
    );
}

#[test]
fn datetime_set_time_timezone_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$utc = new DateTimeZone('UTC');
$dt = new DateTime('2024-01-01 12:00:00', $utc);
$la = new DateTimeZone('America/Los_Angeles');
$dt->setTimezone($la);
echo $dt->format('H');
echo $dt->format('e');
"#
        ),
        &["04America/Los_Angeles"]
    );
}

#[test]
fn datetime_timezone_offset_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$tz = new DateTimeZone('Europe/Paris');
$dt = new DateTime('2024-07-01 12:00:00', $tz);
echo $tz->getOffset($dt);
"#
        ),
        &["7200"]
    );
}

#[test]
fn datetime_isodate_and_timestamp_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime('2024-03-08 00:00:00+00:00');
echo $dt->format(DateTime::ATOM), '|';
echo $dt->format('c');
"#
        ),
        &["2024-03-08T00:00:00+00:00|2024-03-08T00:00:00+00:00"]
    );
}

#[test]
fn datetime_timezone_names_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$tz = new DateTimeZone('America/New_York');
$zones = DateTimeZone::listIdentifiers();
$has_ny = in_array('America/New_York', $zones, true) ? 'yes' : 'no';
echo $has_ny . '|' . $tz->getName();
"#
        ),
        &["yes|America/New_York"]
    );
}

#[test]
fn microtime_integer_and_float_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo is_float(microtime(true)) ? 'float' : 'int';
$micro = microtime();
echo str_contains($micro, ' ') ? 'sp' : 'ns';
echo strpos($micro, '.') !== false ? '|dot' : '|nodot';
"#
        ),
        &["floatsp|dot"]
    );
}

#[test]
fn test_date_sun_info_latitude_longitude() {
    assert_eq!(
        run_prints(
            r#"<?php
$sun = date_sun_info(mktime(0, 0, 0, 6, 21, 2024), 51.5074, -0.1278);
echo is_array($sun) && isset($sun['sunrise']) ? 'sun_info_ok' : 'err';
"#
        ),
        &["sun_info_ok"]
    );
}

#[test]
fn test_date_create_procedural_alias() {
    assert_eq!(
        run_prints(
            r#"<?php
$d = date_create('2024-11-05');
echo date_format($d, 'Y-m-d');
"#
        ),
        &["2024-11-05"]
    );
}

#[test]
fn test_date_parse_strict_and_error_state() {
    assert_eq!(
        run_prints(
            r#"<?php
$d = date_parse('2024-12-01 12:00:00');
echo is_array($d) ? 'ok' : 'bad';
echo $d['error_count'];
echo '|';
$bad = date_parse('2024-13-99');
echo is_array($bad['errors']) ? 'errors' : 'noerr';
"#
        ),
        &["ok0|errors"]
    );
}

#[test]
fn test_strtotime_with_timezone_name() {
    assert_eq!(
        run_prints(
            r#"<?php
$base = strtotime('2024-01-01 00:00:00 UTC');
echo date('Y-m-d', $base);
echo '|';
$z = new DateTimeZone('UTC');
echo $z->getName();
"#
        ),
        &["2024-01-01|UTC"]
    );
}

#[test]
fn test_datetime_interval_spec_years_months() {
    assert_eq!(
        run_prints(
            r#"<?php
$i = new DateInterval('P1Y2M10D');
echo $i->y;
echo $i->m;
echo $i->d;
"#
        ),
        &["1210"]
    );
}

#[test]
fn test_timezone_constructor_and_offset_change() {
    assert_eq!(
        run_prints(
            r#"<?php
$tz = new DateTimeZone('Europe/Paris');
$offset = $tz->getOffset(new DateTime('2024-07-01', $tz));
echo $offset > 3000 ? 'summer' : 'winter';
echo '|';
echo $tz->getName();
"#
        ),
        &["summer|Europe/Paris"]
    );
}

#[test]
fn test_datetime_immutable_modify_does_not_mutate_original() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = new DateTimeImmutable('2024-01-01');
$b = $a->modify('+10 days');
echo $a->format('Y-m-d');
echo $b->format('Y-m-d');
"#
        ),
        &["2024-01-012024-01-11"]
    );
}

#[test]
fn datetime_utc_and_non_utc_parse_behavior() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = new DateTime('2024-01-01 00:00:00+00:00');
$b = new DateTime('2024-01-01 00:00:00-05:00');
echo $a == $b ? 'equal' : 'diff';
echo $a < $b ? '|before' : '|notbefore';
"#
        ),
        &["diff|before"]
    );
}

#[test]
fn test_datetime_timezone_transition_list() {
    assert_eq!(
        run_prints(
            r#"<?php
$tz = new DateTimeZone('America/New_York');
$transitions = $tz->getTransitions(strtotime('2024-01-01'), strtotime('2024-12-31'));
echo is_array($transitions) ? 'arr' : 'na';
echo '|';
echo count($transitions) > 0 ? 'many' : 'none';
"#
        ),
        &["arr|many"]
    );
}

#[test]
fn datetime_create_from_format_with_timezone_designator() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = DateTime::createFromFormat('Y-m-d H:i:s P', '2024-08-01 10:15:00 +02:00');
echo $dt !== false ? 'ok' : 'bad';
echo $dt ? '|' . $dt->format('P') : '|err';
"#
        ),
        &["ok|+02:00"]
    );
}

#[test]
fn datetime_modify_then_clone_preserves_original() {
    assert_eq!(
        run_prints(
            r#"<?php
$base = new DateTime('2024-01-01 00:00:00');
$clone = clone $base;
$base->modify('+1 day');
echo $base->format('Y-m-d');
echo '|';
echo $clone->format('Y-m-d');
"#
        ),
        &["2024-01-02|2024-01-01"]
    );
}

#[test]
fn datetime_parsing_fractional_seconds_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime('2024-01-01 12:34:56.123456');
$fmt = $dt->format('H:i:s.u');
echo str_contains($fmt, '.') ? 'dot' : 'n';
echo '|';
echo is_numeric(str_replace('.', '', explode('.', $fmt)[1])) ? 'micro' : 'no';
"#
        ),
        &["dot|micro"]
    );
}

#[test]
fn datetime_timezone_abbreviations_snapshot() {
    assert_eq!(
        run_prints(
            r#"<?php
$abbr = DateTimeZone::listAbbreviations();
echo is_array($abbr) ? 'arr' : 'na';
echo '|';
echo isset($abbr['est']) ? 'has' : 'missing';
"#
        ),
        &["arr|has"]
    );
}

#[test]
fn datetime_timezone_has_transitions_in_range() {
    assert_eq!(
        run_prints(
            r#"<?php
$tz = new DateTimeZone('America/Chicago');
$changes = $tz->getLocation();
echo is_array($changes) ? 'arr' : 'na';
echo '|';
echo isset($changes['country_code']) ? 'cc' : 'nocc';
"#
        ),
        &["arr|cc"]
    );
}

#[test]
fn datetime_set_date_drops_time_of_day_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime('2024-01-01 12:34:56');
$dt->setDate(2024, 2, 15);
echo $dt->format('Y-m-d H:i:s');
"#
        ),
        &["2024-02-15 12:34:56"]
    );
}

#[test]
fn datetime_set_time_updates_time_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime('2024-01-01 00:00:00');
$dt->setTime(23, 45, 59);
echo $dt->format('H:i:s');
"#
        ),
        &["23:45:59"]
    );
}

#[test]
fn datetimeparse_invalid_string_is_false() {
    assert_eq!(
        run_prints(
            r#"<?php
$ts = strtotime('not-a-date');
echo $ts === false ? 'no' : 'yes';
"#
        ),
        &["no"]
    );
}

#[test]
fn datetime_default_timezone_snapshot() {
    assert_eq!(
        run_prints(
            r#"<?php
date_default_timezone_set('UTC');
echo date_default_timezone_get();
echo '|';
echo date('Y', 0);
"#
        ),
        &["UTC|1970"]
    );
}

#[test]
fn datetime_gmdate_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
date_default_timezone_set('America/New_York');
echo gmdate('Y-m-d', 0);
"#
        ),
        &["1970-01-01"]
    );
}

#[test]
fn datetime_localtime_fields_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$ts = mktime(13, 5, 7, 2, 3, 2024);
$lt = localtime($ts, true);
echo $lt['tm_mday'];
echo $lt['tm_mon'];
echo $lt['tm_year'] + 1900;
"#
        ),
        &["32024"]
    );
}

#[test]
fn datetime_set_timezone_affects_date_output_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
date_default_timezone_set('UTC');
$utc = new DateTime('2024-01-01 12:00:00');
date_default_timezone_set('America/Los_Angeles');
$la = new DateTime('2024-01-01 12:00:00');
echo $utc->format('Y-m-d H:i');
echo '|';
echo $la->format('Y-m-d H:i');
"#
        ),
        &["2024-01-01 12:00|2024-01-01 12:00"]
    );
}

#[test]
fn datetime_from_string_with_timezone_offset_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime('2024-05-01T08:00:00+05:30');
echo $dt->format('P');
echo '|';
echo $dt->format('H:i');
"#
        ),
        &["+05:30|08:00"]
    );
}

#[test]
fn datetime_interval_hours_minutes_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$i = new DateInterval('PT5H45M30S');
echo $i->h;
echo $i->i;
echo $i->s;
"#
        ),
        &["54530"]
    );
}

#[test]
fn datetime_timezone_transition_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$tz = new DateTimeZone('America/New_York');
$trans = $tz->getTransitions(strtotime('2024-01-01'), strtotime('2024-12-31'));
echo is_array($trans) ? 'arr' : 'na';
echo '|';
echo count($trans) > 1 ? 'many' : 'few';
"#
        ),
        &["arr|many"]
    );
}

#[test]
fn datetime_modify_with_keywords_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime('2024-01-31');
$dt->modify('first day of next month');
echo $dt->format('Y-m-d');
echo '|';
$dt->modify('last day of last month');
echo $dt->format('Y-m-d');
"#
        ),
        &["2024-02-01|2024-01-31"]
    );
}

#[test]
fn datetime_parse_and_checkruntime() {
    assert_eq!(
        run_prints(
            r#"<?php
$p = date_parse('2024-11-30 10:20:30');
echo is_array($p) ? 'ok' : 'bad';
echo '|';
echo isset($p['warning_count']) ? 'warn' . $p['warning_count'] : 'nowarn';
"#
        ),
        &["ok|warn0"]
    );
}

#[test]
fn date_time_create_and_clone_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime('2024-01-01 08:00:00');
$clone = clone $dt;
$dt->modify('+1 day');
echo $dt->format('Y-m-d') . '|' . $clone->format('Y-m-d');
"#,
        ),
        &["2024-01-02|2024-01-01"]
    );
}

#[test]
fn date_time_immutable_chain_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTimeImmutable('2024-01-31 10:00:00', new DateTimeZone('UTC'));
$next = $dt->modify('+1 month')->modify('+1 day');
echo $dt->format('Y-m-d');
echo '|';
echo $next->format('Y-m-d');
"#,
        ),
        &["2024-01-31|2024-03-03"]
    );
}

#[test]
fn date_create_from_format_strict_timezone_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = DateTime::createFromFormat('Y-m-d\TH:i:sP', '2024-06-15T10:20:30-04:00');
echo $dt ? 'ok' : 'bad';
echo '|';
echo $dt->format('P');
"#,
        ),
        &["ok|-04:00"]
    );
}

#[test]
fn date_interval_string_roundtrip_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$i = new DateInterval('P1DT2H30M15S');
echo $i->d;
echo $i->h;
echo $i->i;
echo $i->s;
"#,
        ),
        &["12315"]
    );
}

#[test]
fn date_timezone_abbreviation_lookup_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$tz = new DateTimeZone('America/New_York');
$abbr = $tz->getName();
$time = new DateTime('2024-07-01 12:00:00', $tz);
echo $abbr . '|' . $time->format('e') . '|' . $time->getOffset();
"#,
        ),
        &["America/New_York|America/New_York|-14400"]
    );
}

#[test]
fn date_microtime_now_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$parts = explode(' ', microtime());
echo is_numeric($parts[0]) ? 'num' : 'na';
echo '|';
echo (int)$parts[1] > 0 ? 'ts' : 'nt';
"#,
        ),
        &["num|ts"]
    );
}
