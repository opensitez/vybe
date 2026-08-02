<?php
// vybe-test: php/late_static_binding/lsb_factory_named_constructor
// origin: languages/php/tests/php/test_late_static_binding.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

class Model {
    protected string $type;
    public static function make(): static {
        $obj = new static();
        $obj->type = static::class;
        return $obj;
    }
    public function getType(): string { return $this->type; }
}
class User extends Model {}
class Post extends Model {}
echo User::make()->getType() . ',' . Post::make()->getType();

__vybe_check(ob_get_clean(), "User,Post");
