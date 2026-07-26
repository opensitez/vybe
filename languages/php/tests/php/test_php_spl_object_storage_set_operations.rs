use super::helpers::run_prints;

#[test]
fn test_spl_object_storage_add_all() {
    assert_eq!(
        run_prints(
            r#"<?php
$s1 = new SplObjectStorage();
$s2 = new SplObjectStorage();
$o1 = new stdClass();
$o2 = new stdClass();
$s1->attach($o1);
$s2->attach($o2);

$s1->addAll($s2);
echo $s1->count(), "\n";
"#
        ),
        vec!["2"]
    );
}

#[test]
fn test_spl_object_storage_remove_all() {
    assert_eq!(
        run_prints(
            r#"<?php
$s1 = new SplObjectStorage();
$s2 = new SplObjectStorage();
$o1 = new stdClass();
$o2 = new stdClass();
$s1->attach($o1);
$s1->attach($o2);
$s2->attach($o1);

$s1->removeAll($s2);
echo $s1->count() . ':' . ($s1->contains($o2) ? 'o2_kept' : 'none'), "\n";
"#
        ),
        vec!["1:o2_kept"]
    );
}

#[test]
fn test_spl_object_storage_remove_all_except() {
    assert_eq!(
        run_prints(
            r#"<?php
$s1 = new SplObjectStorage();
$s2 = new SplObjectStorage();
$o1 = new stdClass();
$o2 = new stdClass();
$o3 = new stdClass();
$s1->attach($o1);
$s1->attach($o2);
$s1->attach($o3);

$s2->attach($o1);
$s2->attach($o3);

$s1->removeAllExcept($s2);
echo $s1->count() . ':' . ($s1->contains($o1) && $s1->contains($o3) ? 'intersection_kept' : 'err'), "\n";
"#
        ),
        vec!["2:intersection_kept"]
    );
}

#[test]
fn test_spl_object_storage_get_hash() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = new SplObjectStorage();
$obj = new stdClass();
$hash = $s->getHash($obj);
echo (strlen($hash) > 0 && $hash === spl_object_hash($obj)) ? 'hash_ok' : 'hash_err', "\n";
"#
        ),
        vec!["hash_ok"]
    );
}

#[test]
fn test_spl_object_storage_array_access_offset_exists() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = new SplObjectStorage();
$obj = new stdClass();
$s[$obj] = "associated_data";
echo isset($s[$obj]) ? $s[$obj] : 'missing', "\n";
"#
        ),
        vec!["associated_data"]
    );
}
