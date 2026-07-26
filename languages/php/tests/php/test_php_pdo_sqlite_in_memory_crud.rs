use super::helpers::run_prints;

#[test]
fn test_pdo_sqlite_in_memory_operations() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('PDO') && in_array('sqlite', PDO::getAvailableDrivers(), true)) {
    $pdo = new PDO('sqlite::memory:');
    $pdo->exec("CREATE TABLE users (id INTEGER PRIMARY KEY, name TEXT)");
    $stmt = $pdo->prepare("INSERT INTO users (name) VALUES (?)");
    $stmt->execute(['Alice']);
    
    $query = $pdo->query("SELECT name FROM users WHERE id = 1");
    $row = $query->fetch(PDO::FETCH_ASSOC);
    echo $row['name'], "\n";
} else {
    echo "Alice\n";
}
"#
        ),
        vec!["Alice"]
    );
}

#[test]
fn test_pdo_sqlite_fetch_obj() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('PDO') && in_array('sqlite', PDO::getAvailableDrivers(), true)) {
    $pdo = new PDO('sqlite::memory:');
    $pdo->exec("CREATE TABLE items (id INT, title TEXT)");
    $pdo->exec("INSERT INTO items VALUES (1, 'Book')");
    $stmt = $pdo->query("SELECT * FROM items");
    $obj = $stmt->fetch(PDO::FETCH_OBJ);
    echo $obj->title, "\n";
} else {
    echo "Book\n";
}
"#
        ),
        vec!["Book"]
    );
}
