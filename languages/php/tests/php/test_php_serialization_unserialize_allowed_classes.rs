use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Serialization & Security — serialize, unserialize, allowed_classes option, __serialize/__unserialize, __sleep/__wakeup
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_serialize_unserialize_primitive_types() {
    let out = run_prints(
        r#"<?php
$data = [
    "int" => 42,
    "float" => 3.14,
    "string" => "hello",
    "bool" => true,
    "null" => null,
    "arr" => [1, 2, 3]
];

$serialized = serialize($data);
$restored = unserialize($serialized);
echo $restored["string"] . " " . $restored["int"] . " arr_count=" . count($restored["arr"]);
"#,
    );
    assert_eq!(out, vec!["hello 42 arr_count=3"]);
}

#[test]
fn test_php_unserialize_allowed_classes_false_security() {
    let out = run_prints(
        r#"<?php
class Dangerous {
    public string $cmd = "rm -rf /";
}

$payload = serialize(new Dangerous());
$restored = unserialize($payload, ["allowed_classes" => false]);
echo is_object($restored) ? get_class($restored) : "NOT_OBJECT";
"#,
    );
    assert_eq!(out, vec!["__PHP_Incomplete_Class"]);
}

#[test]
fn test_php_unserialize_allowed_classes_whitelist() {
    let out = run_prints(
        r#"<?php
class AllowedDto {
    public string $name = "Alice";
}

class BlockedDto {
    public string $secret = "123";
}

$p1 = serialize(new AllowedDto());
$p2 = serialize(new BlockedDto());

$r1 = unserialize($p1, ["allowed_classes" => [AllowedDto::class]]);
$r2 = unserialize($p2, ["allowed_classes" => [AllowedDto::class]]);

echo get_class($r1) . " vs " . get_class($r2);
"#,
    );
    assert_eq!(out, vec!["AllowedDto vs __PHP_Incomplete_Class"]);
}

#[test]
fn test_php_php74_custom_serialize_unserialize_methods() {
    let out = run_prints(
        r#"<?php
class UserRecord {
    public function __construct(
        public int $id,
        public string $username,
        public string $passwordHash
    ) {}

    public function __serialize(): array {
        return ["i" => $this->id, "u" => $this->username];
    }

    public function __unserialize(array $data): void {
        $this->id = $data["i"];
        $this->username = $data["u"];
        $this->passwordHash = "";
    }
}

$u = new UserRecord(1, "john_doe", "secret_hash");
$s = serialize($u);
$restored = unserialize($s);
echo "{$restored->id}:{$restored->username} hash=" . strlen($restored->passwordHash);
"#,
    );
    assert_eq!(out, vec!["1:john_doe hash=0"]);
}

#[test]
fn test_php_serialize_stdclass_object() {
    compile_ok(
        r#"<?php
$obj = new stdClass();
$obj->title = "Test";
$obj->tags = ["php", "unit"];

$serialized = serialize($obj);
$restored = unserialize($serialized);
echo $restored->title . " tags=" . count($restored->tags);
"#,
    );
}

#[test]
fn test_php_serialize_enum_cases() {
    compile_ok(
        r#"<?php
enum Role: string { case Admin = "admin"; case User = "user"; }

$s = serialize(Role::Admin);
$restored = unserialize($s);
echo $restored->name . "=" . $restored->value;
"#,
    );
}

#[test]
fn test_php_legacy_sleep_wakeup_serialization() {
    compile_ok(
        r#"<?php
class DbModel {
    public string $table = "users";
    public mixed $connection = "active_res";

    public function __sleep(): array {
        return ["table"];
    }

    public function __wakeup(): void {
        $this->connection = "reconnected";
    }
}

$m = new DbModel();
$s = serialize($m);
$restored = unserialize($s);
echo $restored->connection;
"#,
    );
}

#[test]
fn test_php_unserialize_max_depth_option() {
    compile_ok(
        r#"<?php
$data = [1, [2, [3, [4]]]];
$s = serialize($data);
$restored = unserialize($s, ["max_depth" => 10]);
echo count($restored);
"#,
    );
}

#[test]
fn test_php_serialize_by_reference_object_graph() {
    compile_ok(
        r#"<?php
$parent = new stdClass();
$child = new stdClass();
$parent->child = $child;
$child->parent = $parent;

$s = serialize($parent);
$restored = unserialize($s);
echo get_class($restored->child->parent);
"#,
    );
}

#[test]
fn test_php_unserialize_invalid_string_error() {
    compile_ok(
        r#"<?php
$invalidSerialized = 'a:2:{i:0;s:3:"foo";';
$restored = @unserialize($invalidSerialized);
echo $restored === false ? "UNSERIALIZE_FAILED" : "SUCCESS";
"#,
    );
}
