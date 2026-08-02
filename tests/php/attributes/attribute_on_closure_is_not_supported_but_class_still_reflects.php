<?php
// vybe-test: php/attributes/attribute_on_closure_is_not_supported_but_class_still_reflects
// origin: languages/php/tests/php/test_attributes.rs

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

#[Attribute]
class Note {
    public function __construct(public string $msg) {}
}
#[Note('worker')]
class Worker {
    public function tag(): string {
        return (new ReflectionClass(self::class))->getAttributes(Note::class)[0]->newInstance()->msg;
    }
}
echo (new Worker())->tag();

__vybe_check(ob_get_clean(), "worker");
