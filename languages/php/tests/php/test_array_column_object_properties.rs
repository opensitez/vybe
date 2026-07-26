
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

    array_column_objects_private_prop_access => {
        r#"<?php
class PrivateUser {
    private int $id;
    private string $name;
    public function __construct(int $id, string $name) {
        $this->id = $id;
        $this->name = $name;
    }
}
$users = [new PrivateUser(1, 'Alice'), new PrivateUser(2, 'Bob')];
$ids = array_column($users, 'id');
echo $ids[0] . "|" . $ids[1];
"#,
        ["1|2"]
    };

    array_column_objects_public_map_with_indexing => {
        r#"<?php
class Product {
    public function __construct(public string $sku, public float $price) {}
}
$rows = [
    new Product('X', 10.5),
    new Product('Y', 12.0),
];
$prices = array_column($rows, 'price', 'sku');
echo $prices['X'] . "|" . $prices['Y'];
"#,
        ["10.5|12"]
    };

    array_column_objects_with_null_and_missing => {
        r#"<?php
class SoftUser {
    public function __construct(public int $id, public ?string $name = null) {}
}
$rows = [new SoftUser(1, 'Alice'), new SoftUser(2), new SoftUser(3, 'Bob')];
$names = array_column($rows, 'name');
echo implode('|', $names);
"#,
        ["Alice||Bob"]
    };

    array_column_private_property_access_uses_accessors_or_public_api => {
        r#"<?php
class WithGetter {
    private int $id = 10;
    private string $name = 'N/A';
    public function __get($field) {
        return $field === 'name' ? $this->name : null;
    }
}
$rows = [new WithGetter()];
$names = array_column($rows, 'name');
echo json_encode($names);
"#,
        ["[null]"]
    };

    array_column_with_default_null_index => {
        r#"<?php
class Row {
    public function __construct(public int $id, public ?string $name = null) {}
}
$rows = [new Row(1, 'A'), new Row(2, null)];
$names = array_column($rows, 'name', 'missing');
echo json_encode(array_values($names));
"#,
        ["[null]"]
    };

    array_column_objects_missing_property_returns_null => {
        r#"<?php
class Basic {
    public function __construct(public int $id) {}
}
$rows = [new Basic(1)];
$vals = array_column($rows, 'does_not_exist');
echo $vals[0] === null ? 'null' : 'notnull';
        "#,
        ["null"]
    };

    array_column_objects_non_scalar_index_uses_column_stringification => {
        r#"<?php
class Row {
    public function __construct(public int $id, public string $name) {}
}
$rows = [new Row(1, 'A'), new Row(2, 'B')];
$vals = array_column($rows, 'name', 'id');
echo $vals[1] . '|' . $vals[2];
"#,
        ["A|B"]
    };

    array_column_objects_index_key_collision => {
        r#"<?php
class Row {
    public function __construct(public string $id, public string $name) {}
}
$rows = [new Row('1', 'A'), new Row('01a', 'B'), new Row('1', 'C')];
$vals = array_column($rows, 'name', 'id');
echo $vals['1'] . '|' . $vals['01a'];
"#,
        ["C|B"]
    };

    array_column_objects_accessor_precedence => {
        r#"<?php
class WithBoth {
    private string $name;
    public function __construct(private int $id, string $name) {
        $this->name = $name;
    }
    public function __get($field) {
        return $field === 'name' ? 'magic-' . $this->name : null;
    }
    public function __isset($field): bool {
        return $field === 'name';
    }
}
$rows = [new WithBoth(1, 'X'), new WithBoth(2, 'Y')];
$vals = array_column($rows, 'name');
echo implode('|', $vals);
"#,
        ["magic-X|magic-Y"]
    };

    array_column_objects_public_inheritance => {
        r#"<?php
class Base {
    public function __construct(public int $id) {}
}
class Child extends Base {
    public function __construct(int $id, public string $name) {
        parent::__construct($id);
    }
}
$rows = [new Child(1, 'Ada'), new Child(2, 'Lin')];
$vals = array_column($rows, 'name');
echo implode('|', $vals);
"#,
        ["Ada|Lin"]
    };

    array_column_objects_mixed_input_throws_when_no_array => {
        r#"<?php
	try {
	    $a = array_column('string', 'name');
	    echo 'no-error';
	} catch (Throwable $e) {
    echo 'error';
}
"#,
        ["error"]
    };
}
