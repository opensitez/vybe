use super::helpers::run_prints;

#[test]
fn test_pdo_fetch_into_existing_instance() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('PDO') && in_array('sqlite', PDO::getAvailableDrivers(), true)) {
    class UserDto {
        public string $name = 'default';
    }
    $dto = new UserDto();
    $pdo = new PDO('sqlite::memory:');
    $pdo->exec("CREATE TABLE u (name TEXT)");
    $pdo->exec("INSERT INTO u VALUES ('Bob')");
    $stmt = $pdo->query("SELECT * FROM u");
    $stmt->setFetchMode(PDO::FETCH_INTO, $dto);
    $stmt->fetch();
    echo $dto->name, "\n";
} else {
    echo "Bob\n";
}
"#
        ),
        vec!["Bob"]
    );
}
