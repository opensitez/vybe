use super::helpers::compile_ok;

// ── DatePeriod ────────────────────────────────────────────────

#[test] fn date_period_basic() {
    compile_ok(r#"<?php
$start    = new DateTimeImmutable('2024-01-01');
$interval = new DateInterval('P1M');  // 1 month
$end      = new DateTimeImmutable('2024-06-01');
$period   = new DatePeriod($start, $interval, $end);
$count = 0;
foreach ($period as $dt) { $count++; }
echo $count;  // 5 (Jan through May)
"#);
}

#[test] fn date_period_daily() {
    compile_ok(r#"<?php
$start = new DateTimeImmutable('2024-03-01');
$end   = new DateTimeImmutable('2024-03-08');
$period = new DatePeriod($start, new DateInterval('P1D'), $end);
$dates = [];
foreach ($period as $dt) { $dates[] = $dt->format('d'); }
echo implode(',', $dates);
"#);
}

#[test] fn date_period_weekly() {
    compile_ok(r#"<?php
$start = new DateTimeImmutable('2024-01-01');
$period = new DatePeriod($start, new DateInterval('P1W'), 4);
$weeks = [];
foreach ($period as $dt) { $weeks[] = $dt->format('W'); }
echo count($weeks);
"#);
}

#[test] fn date_period_recurrences() {
    compile_ok(r#"<?php
$start = new DateTimeImmutable('2024-01-15');
$period = new DatePeriod($start, new DateInterval('P1M'), 5);
$months = [];
foreach ($period as $dt) { $months[] = $dt->format('m'); }
echo implode(',', $months);
"#);
}

#[test] fn date_period_include_start_date() {
    compile_ok(r#"<?php
$start = new DateTimeImmutable('2024-01-01');
$end   = new DateTimeImmutable('2024-04-01');
// With DatePeriod::INCLUDE_START_DATE (default) vs EXCLUDE_START_DATE
$incl = new DatePeriod($start, new DateInterval('P1M'), $end);
$count = 0;
foreach ($incl as $dt) { $count++; }
echo $count;  // 3
"#);
}

#[test] fn date_period_get_start_end() {
    compile_ok(r#"<?php
$start = new DateTimeImmutable('2024-01-01');
$end   = new DateTimeImmutable('2024-12-31');
$period = new DatePeriod($start, new DateInterval('P1M'), $end);
echo $period->getStartDate()->format('Y-m-d');
echo ':' . $period->getEndDate()->format('Y-m-d');
"#);
}

#[test] fn date_period_get_date_interval() {
    compile_ok(r#"<?php
$start = new DateTimeImmutable('2024-01-01');
$interval = new DateInterval('P2W');
$period = new DatePeriod($start, $interval, 3);
echo $period->getDateInterval()->days >= 14 ? 'interval ok' : 'wrong';
"#);
}

// ── DateInterval deep ─────────────────────────────────────────

#[test] fn date_interval_create_from_date_string() {
    compile_ok(r#"<?php
$i = DateInterval::createFromDateString('2 weeks + 3 days');
echo $i->days >= 0 ? 'created' : 'failed';
"#);
}

#[test] fn date_interval_format() {
    compile_ok(r#"<?php
$i = new DateInterval('P1Y2M3DT4H5M6S');
echo $i->format('%Y years, %M months, %D days');
echo ':' . $i->format('%H:%I:%S');
"#);
}

#[test] fn date_interval_invert() {
    compile_ok(r#"<?php
$a = new DateTimeImmutable('2024-03-01');
$b = new DateTimeImmutable('2024-01-01');
$diff = $a->diff($b);
echo $diff->invert;  // 1 (b < a, so diff is negative direction)
echo ':' . $diff->m;
"#);
}

#[test] fn date_interval_components() {
    compile_ok(r#"<?php
$i = new DateInterval('P3Y6M15DT12H30M45S');
echo $i->y . ':' . $i->m . ':' . $i->d;
echo ':' . $i->h . ':' . $i->i . ':' . $i->s;
"#);
}

#[test] fn date_diff_abs_days() {
    compile_ok(r#"<?php
$a = new DateTimeImmutable('2024-01-01');
$b = new DateTimeImmutable('2024-12-31');
$diff = $a->diff($b);
echo $diff->days;  // 365 (2024 is leap year)
"#);
}

// ── DateTimeZone ──────────────────────────────────────────────

#[test] fn datetime_timezone_basic() {
    compile_ok(r#"<?php
$tz = new DateTimeZone('America/New_York');
echo $tz->getName();
"#);
}

#[test] fn datetime_timezone_offset() {
    compile_ok(r#"<?php
$tz  = new DateTimeZone('UTC');
$dt  = new DateTime('now', $tz);
$off = $tz->getOffset($dt);
echo $off;  // 0 for UTC
"#);
}

#[test] fn datetime_timezone_list_by_region() {
    compile_ok(r#"<?php
$zones = DateTimeZone::listIdentifiers(DateTimeZone::EUROPE);
echo count($zones) > 0 ? 'has zones' : 'empty';
echo in_array('Europe/London', $zones) ? ':has London' : ':no London';
"#);
}

#[test] fn datetime_timezone_all_identifiers() {
    compile_ok(r#"<?php
$all = DateTimeZone::listIdentifiers();
echo count($all) > 400 ? 'many zones' : 'few zones';
"#);
}

#[test] fn datetime_convert_timezone() {
    compile_ok(r#"<?php
$utc = new DateTimeImmutable('2024-06-15 12:00:00', new DateTimeZone('UTC'));
$ny  = $utc->setTimezone(new DateTimeZone('America/New_York'));
$tok = $utc->setTimezone(new DateTimeZone('Asia/Tokyo'));
// NY is UTC-4 in summer, Tokyo is UTC+9
echo $ny->format('H')  . ':' . $tok->format('H');
"#);
}

#[test] fn datetime_timezone_transitions() {
    compile_ok(r#"<?php
$tz = new DateTimeZone('America/New_York');
$transitions = $tz->getTransitions(mktime(0,0,0,1,1,2024), mktime(0,0,0,12,31,2024));
echo count($transitions) > 0 ? 'has transitions' : 'no transitions';
"#);
}

// ── DateTime / DateTimeImmutable patterns ─────────────────────

#[test] fn datetime_immutable_modify() {
    compile_ok(r#"<?php
$dt = new DateTimeImmutable('2024-01-15');
$next_month = $dt->modify('+1 month');
echo $dt->format('Y-m-d');          // unchanged
echo ':' . $next_month->format('Y-m-d');
"#);
}

#[test] fn datetime_add_sub() {
    compile_ok(r#"<?php
$dt = new DateTimeImmutable('2024-01-01');
$plus30 = $dt->add(new DateInterval('P30D'));
$minus7 = $dt->sub(new DateInterval('P7D'));
echo $plus30->format('Y-m-d');
echo ':' . $minus7->format('Y-m-d');
"#);
}

#[test] fn datetime_create_from_format() {
    compile_ok(r#"<?php
$dt = DateTime::createFromFormat('d/m/Y H:i', '15/06/2024 14:30');
echo $dt->format('Y-m-d H:i');
$dt2 = DateTimeImmutable::createFromFormat('U', '1718438400');
echo ':' . ($dt2 !== false ? 'ok' : 'fail');
"#);
}

#[test] fn datetime_timestamp() {
    compile_ok(r#"<?php
$dt = new DateTimeImmutable('2024-01-01 00:00:00', new DateTimeZone('UTC'));
$ts = $dt->getTimestamp();
echo $ts > 0 ? 'positive ts' : 'non-positive';
$back = (new DateTimeImmutable())->setTimestamp($ts);
echo ':' . $back->format('Y');
"#);
}

#[test] fn datetime_comparison() {
    compile_ok(r#"<?php
$a = new DateTimeImmutable('2024-01-01');
$b = new DateTimeImmutable('2024-06-15');
echo ($a < $b)  ? 'a before b' : 'a not before b';
echo ($a == $b) ? ':equal'     : ':not equal';
"#);
}

#[test] fn datetime_create_from_interface() {
    compile_ok(r#"<?php
$mutable = new DateTime('2024-03-15');
$immutable = DateTimeImmutable::createFromMutable($mutable);
echo $immutable->format('Y-m-d');
echo ($immutable instanceof DateTimeImmutable) ? ':immutable' : ':not immutable';
"#);
}

// ── date_sun_info ─────────────────────────────────────────────

#[test] fn date_sun_info_basic() {
    compile_ok(r#"<?php
$info = date_sun_info(mktime(0,0,0,6,21,2024), 51.5, -0.12); // London midsummer
echo isset($info['sunrise'])  ? 'has sunrise' : 'no sunrise';
echo isset($info['sunset'])   ? ':has sunset' : ':no sunset';
echo isset($info['transit'])  ? ':has transit' : ':no transit';
"#);
}

// ── Practical date patterns ───────────────────────────────────

#[test] fn business_days_between() {
    compile_ok(r#"<?php
function businessDays(DateTimeImmutable $start, DateTimeImmutable $end): int {
    $count = 0;
    $period = new DatePeriod($start, new DateInterval('P1D'), $end);
    foreach ($period as $day) {
        $dow = (int)$day->format('N'); // 1=Mon ... 7=Sun
        if ($dow < 6) $count++;
    }
    return $count;
}
$start = new DateTimeImmutable('2024-01-01');
$end   = new DateTimeImmutable('2024-01-08');
echo businessDays($start, $end);
"#);
}

#[test] fn month_boundaries() {
    compile_ok(r#"<?php
$months = [];
$start = new DateTimeImmutable('2024-01-01');
$period = new DatePeriod($start, new DateInterval('P1M'), 12);
foreach ($period as $dt) {
    $months[] = $dt->format('Y-m');
}
echo count($months) . ':' . $months[0] . ':' . end($months);
"#);
}

#[test] fn age_calculation() {
    compile_ok(r#"<?php
function calculateAge(DateTimeImmutable $birthday, DateTimeImmutable $today): int {
    return $birthday->diff($today)->y;
}
$birthday = new DateTimeImmutable('1990-06-15');
$today    = new DateTimeImmutable('2024-06-15');
echo calculateAge($birthday, $today);
"#);
}

#[test] fn recurring_monthly_dates() {
    compile_ok(r#"<?php
$start = new DateTimeImmutable('2024-01-31');
$dates = [];
for ($i = 0; $i < 4; $i++) {
    $dates[] = $start->modify("+$i month")->format('Y-m-d');
}
echo implode(',', $dates);
"#);
}
