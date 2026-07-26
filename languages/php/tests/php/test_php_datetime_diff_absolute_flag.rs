use super::helpers::run_prints;

#[test]
fn test_datetime_diff_absolute_true() {
    assert_eq!(
        run_prints(
            r#"<?php
$d1 = new DateTime('2024-01-01');
$d2 = new DateTime('2024-01-10');
$diff = $d1->diff($d2, true);
echo $diff->days . ':' . ($diff->invert ? 'neg' : 'pos'), "\n";
"#
        ),
        vec!["9:pos"]
    );
}

#[test]
fn test_datetime_diff_invert_flag() {
    assert_eq!(
        run_prints(
            r#"<?php
$d1 = new DateTime('2024-01-10');
$d2 = new DateTime('2024-01-01');
$diff = $d1->diff($d2, false);
echo $diff->days . ':' . ($diff->invert ? '1' : '0'), "\n";
"#
        ),
        vec!["9:1"]
    );
}

#[test]
fn test_datetime_diff_absolute_true_zero_days() {
    assert_eq!(
        run_prints(
            r#"<?php
$d1 = new DateTime('2024-01-01 12:00:00');
$d2 = new DateTime('2024-01-01 12:00:00');
$diff = $d1->diff($d2, true);
echo $diff->days . ':' . ($diff->invert ? 'neg' : 'pos');
"#
        ),
        vec!["0:pos"]
    );
}

#[test]
fn test_datetime_diff_absolute_false_same_invert() {
    assert_eq!(
        run_prints(
            r#"<?php
$d1 = new DateTime('2024-06-01 00:00:00');
$d2 = new DateTime('2024-05-01 00:00:00');
$diff = $d1->diff($d2, false);
echo $diff->days . ':' . $diff->invert;
"#
        ),
        vec!["31:1"]
    );
}

#[test]
fn test_datetime_diff_absolute_tz_gap() {
    assert_eq!(
        run_prints(
            r#"<?php
$d1 = new DateTime('2024-01-01 00:00:00', new DateTimeZone('UTC'));
$d2 = new DateTime('2024-01-01 00:00:00', new DateTimeZone('Asia/Tokyo'));
$diff = $d1->diff($d2, true);
echo $diff->days . ':' . ($diff->invert ? 'neg' : 'pos');
"#
        ),
        vec!["0:pos"]
    );
}

#[test]
fn test_datetime_diff_absolute_end_of_month_edge() {
    assert_eq!(
        run_prints(
            r#"<?php
$d1 = new DateTime('2024-01-31');
$d2 = new DateTime('2024-02-29');
$diff = $d1->diff($d2, true);
echo $diff->days . ':' . $diff->m . ':' . $diff->d;
"#
        ),
        vec!["29:0:29"]
    );
}
