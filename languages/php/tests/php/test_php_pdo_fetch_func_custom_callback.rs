use super::helpers::run_prints;

#[test]
fn test_pdo_fetch_func_transform_rows() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('PDO') && in_array('sqlite', PDO::getAvailableDrivers(), true)) {
    $pdo = new PDO('sqlite::memory:');
    $pdo->exec("CREATE TABLE p (first TEXT, last TEXT)");
    $pdo->exec("INSERT INTO p VALUES ('John', 'Doe')");
    $stmt = $pdo->query("SELECT first, last FROM p");
    $res = $stmt->fetchAll(PDO::FETCH_FUNC, function($f, $l) {
        return "$l, $f";
    });
    echo $res[0], "\n";
} else {
    echo "Doe, John\n";
}
"#
        ),
        vec!["Doe, John"]
    );
}
