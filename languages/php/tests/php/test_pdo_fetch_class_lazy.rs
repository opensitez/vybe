
crate::php_cases! {
    pdo_fetch_class_instantiation => {
        r#"<?php
class User {
    public string $name;
    public int $age;
    
    public function __construct() {
        // PDO assigns properties BEFORE calling __construct by default
        echo $this->name . ':';
    }
}

$pdo = new PDO('sqlite::memory:');
$pdo->exec("CREATE TABLE users (name TEXT, age INTEGER)");
$pdo->exec("INSERT INTO users VALUES ('Alice', 30)");

$stmt = $pdo->query("SELECT name, age FROM users");
$user = $stmt->fetchObject(User::class);
echo $user->age;
"#,
        ["Alice:30"]
    };

    pdo_fetch_class_late_props => {
        r#"<?php
class UserLate {
    public string $name;
}

$pdo = new PDO('sqlite::memory:');
$pdo->exec("CREATE TABLE users (name TEXT)");
$pdo->exec("INSERT INTO users VALUES ('Bob')");

$stmt = $pdo->query("SELECT name FROM users");
$stmt->setFetchMode(PDO::FETCH_CLASS | PDO::FETCH_PROPS_LATE, UserLate::class);
$user = $stmt->fetch();
echo $user->name;
"#,
        ["Bob"]
    };
}
