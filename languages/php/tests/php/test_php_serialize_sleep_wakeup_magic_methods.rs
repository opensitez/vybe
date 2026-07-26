use super::helpers::run_prints;

#[test]
fn test_serialize_sleep_and_wakeup_hooks() {
    assert_eq!(
        run_prints(
            r#"<?php
class SerializableObj {
    public string $keep = 'saved';
    public string $discard = 'ignored';
    public bool $restored = false;

    public function __sleep(): array {
        return ['keep'];
    }

    public function __wakeup(): void {
        $this->restored = true;
    }
}

$s = serialize(new SerializableObj());
$obj = unserialize($s);
echo $obj->keep . '|' . ($obj->restored ? 'restored' : 'not_restored'), "\n";
"#
        ),
        vec!["saved|restored"]
    );
}
