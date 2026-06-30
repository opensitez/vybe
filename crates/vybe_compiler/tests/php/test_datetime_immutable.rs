//! `DateTimeImmutable` and immutable date math (UTC-fixed expectations).

crate::php_cases! {
    datetime_immutable_create_from_format => {
        r#"<?php
date_default_timezone_set('UTC');
$d = DateTimeImmutable::createFromFormat('Y-m-d', '2024-06-01');
echo $d->format('Y-m-d');
"#,
        ["2024-06-01"]
    };

    datetime_immutable_modify_returns_new_instance => {
        r#"<?php
date_default_timezone_set('UTC');
$a = new DateTimeImmutable('2024-01-01', new DateTimeZone('UTC'));
$b = $a->modify('+1 day');
echo $a->format('d') . $b->format('d');
"#,
        ["102"]
    };

    datetime_immutable_set_timezone => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-01-01 12:00:00', new DateTimeZone('UTC'));
$d2 = $d->setTimezone(new DateTimeZone('Europe/London'));
echo $d2->getTimezone()->getName();
"#,
        ["Europe/London"]
    };

    datetime_immutable_add_interval => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-01-01', new DateTimeZone('UTC'));
$d2 = $d->add(new DateInterval('P2D'));
echo $d2->format('Y-m-d');
"#,
        ["2024-01-03"]
    };

    datetime_immutable_sub_interval => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-01-10', new DateTimeZone('UTC'));
$d2 = $d->sub(new DateInterval('P3D'));
echo $d2->format('d');
"#,
        ["07"]
    };

    datetime_immutable_diff_days => {
        r#"<?php
date_default_timezone_set('UTC');
$a = new DateTimeImmutable('2024-01-01', new DateTimeZone('UTC'));
$b = new DateTimeImmutable('2024-01-04', new DateTimeZone('UTC'));
echo $a->diff($b)->days;
"#,
        ["3"]
    };

    datetime_immutable_get_timestamp => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-01-01 00:00:00', new DateTimeZone('UTC'));
echo $d->getTimestamp() > 0 ? 'pos' : 'zero';
"#,
        ["pos"]
    };

    datetime_immutable_set_date => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-01-01', new DateTimeZone('UTC'));
$d2 = $d->setDate(2025, 12, 25);
echo $d2->format('Y-m-d');
"#,
        ["2025-12-25"]
    };

    datetime_immutable_set_time => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-01-01', new DateTimeZone('UTC'));
$d2 = $d->setTime(14, 30, 0);
echo $d2->format('H:i');
"#,
        ["14:30"]
    };

    datetime_immutable_format_rfc3339 => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-06-15T10:00:00+00:00');
echo $d->format('c');
"#,
        ["2024-06-15T10:00:00+00:00"]
    };

    datetime_immutable_w_format_weekday => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-06-17', new DateTimeZone('UTC'));
echo $d->format('N');
"#,
        ["1"]
    };

    datetime_immutable_original_unchanged_after_ops => {
        r#"<?php
date_default_timezone_set('UTC');
$orig = new DateTimeImmutable('2024-05-01', new DateTimeZone('UTC'));
$orig->add(new DateInterval('P10D'));
echo $orig->format('Y-m-d');
"#,
        ["2024-05-01"]
    };

    datetime_immutable_create_from_interface => {
        r#"<?php
date_default_timezone_set('UTC');
$mutable = new DateTime('2024-03-01', new DateTimeZone('UTC'));
$imm = DateTimeImmutable::createFromMutable($mutable);
echo $imm->format('Y-m-d');
"#,
        ["2024-03-01"]
    };

    datetime_immutable_json_serialize => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-01-02T03:04:05+00:00');
echo json_encode($d);
"#,
        ["\"2024-01-02T03:04:05+00:00\""]
    };

    datetime_immutable_compare_spaceship => {
        r#"<?php
date_default_timezone_set('UTC');
$a = new DateTimeImmutable('2024-01-01', new DateTimeZone('UTC'));
$b = new DateTimeImmutable('2024-01-02', new DateTimeZone('UTC'));
echo $a <=> $b;
"#,
        ["-1"]
    };

    datetime_immutable_get_offset_utc => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-07-01', new DateTimeZone('UTC'));
echo $d->getOffset();
"#,
        ["0"]
    };

    datetime_immutable_last_day_of_month => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-02-01', new DateTimeZone('UTC'));
$d2 = $d->modify('last day of this month');
echo $d2->format('d');
"#,
        ["29"]
    };

    datetime_immutable_first_day_of_year => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-06-15', new DateTimeZone('UTC'));
$d2 = $d->modify('first day of january');
echo $d2->format('m-d');
"#,
        ["01-01"]
    };

    datetime_immutable_add_months => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-01-31', new DateTimeZone('UTC'));
$d2 = $d->add(new DateInterval('P1M'));
echo $d2->format('Y-m-d');
"#,
        ["2024-03-02"]
    };

    datetime_immutable_sub_hours => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-01-01 05:00:00', new DateTimeZone('UTC'));
$d2 = $d->sub(new DateInterval('PT2H'));
echo $d2->format('H');
"#,
        ["03"]
    };

    datetime_immutable_timezone_get => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('now', new DateTimeZone('UTC'));
echo $d->getTimezone()->getName();
"#,
        ["UTC"]
    };

    datetime_immutable_create_from_timestamp => {
        r#"<?php
date_default_timezone_set('UTC');
$d = (new DateTimeImmutable('@0'))->setTimezone(new DateTimeZone('UTC'));
echo $d->format('Y');
"#,
        ["1970"]
    };

    datetime_immutable_week_number => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-01-08', new DateTimeZone('UTC'));
echo $d->format('W');
"#,
        ["02"]
    };

    datetime_immutable_day_of_year => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-12-31', new DateTimeZone('UTC'));
echo $d->format('z');
"#,
        ["365"]
    };

    datetime_immutable_is_leap_year => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-02-29', new DateTimeZone('UTC'));
echo $d->format('L');
"#,
        ["1"]
    };

    datetime_immutable_non_leap_feb => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2023-02-28', new DateTimeZone('UTC'));
$d2 = $d->modify('+1 day');
echo $d2->format('m-d');
"#,
        ["03-01"]
    };

    datetime_immutable_microseconds_in_format => {
        r#"<?php
date_default_timezone_set('UTC');
$d = DateTimeImmutable::createFromFormat('Y-m-d H:i:s.u', '2024-01-01 00:00:00.123456', new DateTimeZone('UTC'));
echo substr($d->format('u'), 0, 3);
"#,
        ["123"]
    };

    datetime_immutable_relative_next_monday => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-06-14', new DateTimeZone('UTC'));
$d2 = $d->modify('next monday');
echo $d2->format('N');
"#,
        ["1"]
    };

    datetime_immutable_state_after_set_is_new => {
        r#"<?php
date_default_timezone_set('UTC');
$a = new DateTimeImmutable('2024-01-01', new DateTimeZone('UTC'));
$b = $a->setDate(2099, 1, 1);
echo ($a->format('Y') === '2024' && $b->format('Y') === '2099') ? 'split' : 'same';
"#,
        ["split"]
    };

    datetime_immutable_parse_atom => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-06-01T00:00:00+00:00');
echo $d->format('Y-m-d');
"#,
        ["2024-06-01"]
    };

    datetime_immutable_diff_invert_flag => {
        r#"<?php
date_default_timezone_set('UTC');
$a = new DateTimeImmutable('2024-01-10', new DateTimeZone('UTC'));
$b = new DateTimeImmutable('2024-01-01', new DateTimeZone('UTC'));
echo $a->diff($b)->invert;
"#,
        ["1"]
    };

    datetime_immutable_format_24h => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-01-01 23:59:59', new DateTimeZone('UTC'));
echo $d->format('H:i:s');
"#,
        ["23:59:59"]
    };

    datetime_immutable_iso_year => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-12-30', new DateTimeZone('UTC'));
echo $d->format('o');
"#,
        ["2025"]
    };

    datetime_immutable_month_name => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-03-15', new DateTimeZone('UTC'));
echo $d->format('F');
"#,
        ["March"]
    };

    datetime_immutable_short_month => {
        r#"<?php
date_default_timezone_set('UTC');
$d = new DateTimeImmutable('2024-03-15', new DateTimeZone('UTC'));
echo $d->format('M');
"#,
        ["Mar"]
    };
}
