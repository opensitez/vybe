use super::helpers::run_prints;

#[test]
fn test_datetime_create_from_interface_immutable_to_mutable() {
    assert_eq!(
        run_prints(
            r#"<?php
$imm = new DateTimeImmutable('2024-05-15 10:30:00', new DateTimeZone('UTC'));
$mut = DateTime::createFromInterface($imm);
echo $mut->format('Y-m-d H:i:s') . ':' . get_class($mut), "\n";
"#
        ),
        vec!["2024-05-15 10:30:00:DateTime"]
    );
}

#[test]
fn test_datetime_create_from_interface_mutable_to_immutable() {
    assert_eq!(
        run_prints(
            r#"<?php
$mut = new DateTime('2024-12-31 23:59:59', new DateTimeZone('UTC'));
$imm = DateTimeImmutable::createFromInterface($mut);
echo $imm->format('Y-m-d H:i:s') . ':' . get_class($imm), "\n";
"#
        ),
        vec!["2024-12-31 23:59:59:DateTimeImmutable"]
    );
}

#[test]
fn test_datetime_create_from_interface_preserves_timezone() {
    assert_eq!(
        run_prints(
            r#"<?php
$tz = new DateTimeZone('Europe/Paris');
$imm = new DateTimeImmutable('2024-07-01 12:00:00', $tz);
$mut = DateTime::createFromInterface($imm);
echo $mut->getTimezone()->getName(), "\n";
"#
        ),
        vec!["Europe/Paris"]
    );
}

#[test]
fn test_datetime_create_from_interface_mutation_independence() {
    assert_eq!(
        run_prints(
            r#"<?php
$imm = new DateTimeImmutable('2024-01-01');
$mut = DateTime::createFromInterface($imm);
$mut->modify('+1 day');
echo $imm->format('Y-m-d') . ',' . $mut->format('Y-m-d'), "\n";
"#
        ),
        vec!["2024-01-01,2024-01-02"]
    );
}
