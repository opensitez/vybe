use super::helpers::run_prints;

#[test]
fn test_datetime_modify_first_day_of_month() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime('2024-03-15', new DateTimeZone('UTC'));
$dt->modify('first day of this month');
echo $dt->format('Y-m-d'), "\n";
"#
        ),
        vec!["2024-03-01"]
    );
}

#[test]
fn test_datetime_modify_last_day_of_next_month() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime('2024-01-10', new DateTimeZone('UTC'));
$dt->modify('last day of next month');
echo $dt->format('Y-m-d'), "\n";
"#
        ),
        vec!["2024-02-29"]
    );
}

#[test]
fn test_datetime_modify_first_day_of_next_year() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime('2024-07-19', new DateTimeZone('UTC'));
$dt->modify('first day of next year');
echo $dt->format('Y-m-d'), "\n";
"#
        ),
        vec!["2025-01-01"]
    );
}

#[test]
fn test_datetime_modify_last_day_of_last_year() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime('2024-07-19', new DateTimeZone('UTC'));
$dt->modify('last day of last year');
echo $dt->format('Y-m-d'), "\n";
"#
        ),
        vec!["2023-12-31"]
    );
}

#[test]
fn test_datetime_modify_chain_with_timezone() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime('2024-01-31', new DateTimeZone('Europe/Paris'));
$dt->modify('+1 day');
$dt->modify('last day of this month');
echo $dt->format('Y-m-d'), "\n";
"#
        ),
        vec!["2024-01-31"]
    );
}

#[test]
fn test_datetime_modify_invalid_relative_phrase() {
    assert_eq!(
        run_prints(
            r#"<?php
$dt = new DateTime('2024-01-01', new DateTimeZone('UTC'));
$res = $dt->modify('very invalid phrase');
echo $res === false ? 'bad' : 'ok';
echo '|';
echo $dt->format('Y-m-d');
"#
        ),
        vec!["bad|2024-01-01"]
    );
}
