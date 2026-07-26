use super::helpers::run_prints;

#[test]
fn test_get_resource_id_stream() {
    assert_eq!(
        run_prints(
            r#"<?php
$f = fopen('php://memory', 'r+');
$id = get_resource_id($f);
fclose($f);
echo is_int($id) && $id > 0 ? 'id_ok' : 'err', "\n";
"#
        ),
        vec!["id_ok"]
    );
}

#[test]
fn test_get_resource_type_stream() {
    assert_eq!(
        run_prints(
            r#"<?php
$f = fopen('php://memory', 'r+');
$type = get_resource_type($f);
fclose($f);
echo $type === 'stream' ? 'stream_type' : $type, "\n";
"#
        ),
        vec!["stream_type"]
    );
}

#[test]
fn test_get_resource_type_closed_stream() {
    assert_eq!(
        run_prints(
            r#"<?php
$f = fopen('php://memory', 'r+');
fclose($f);
$type = get_resource_type($f);
echo $type === 'Unknown' || str_contains($type, 'closed') ? 'closed_ok' : 'ok', "\n";
"#
        ),
        vec!["closed_ok"]
    );
}
