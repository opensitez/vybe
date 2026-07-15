use super::helpers::run_prints;

crate::php_cases! {
    array_column_objects_public_props => {
        r#"<?php
class User {
    public function __construct(public int $id, public string $name) {}
}
$users = [
    new User(1, 'Alice'),
    new User(2, 'Bob'),
];
$names = array_column($users, 'name', 'id');
echo $names[1] . "|" . $names[2];
"#,
        ["Alice|Bob"]
    };

    array_column_objects_magic_get => {
        r#"<?php
class MagicUser {
    private $data;
    public function __construct($id, $name) {
        $this->data = ['id' => $id, 'name' => $name];
    }
    public function __isset($name) {
        return isset($this->data[$name]);
    }
    public function __get($name) {
        return $this->data[$name];
    }
}
$users = [
    new MagicUser(10, 'Eve'),
    new MagicUser(20, 'Mallory'),
];
$ids = array_column($users, 'id');
echo implode(',', $ids);
"#,
        ["10,20"]
    };
}
