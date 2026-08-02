<?php
// vybe-test: php/patterns/memento_capture_restore
// origin: languages/php/tests/php/test_patterns.rs

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

class EditorMemento {
    public function __construct(public readonly string $content) {}
}
class Editor {
    public $content = '';
    public function save(): EditorMemento { return new EditorMemento($this->content); }
    public function restore(EditorMemento $m): void { $this->content = $m->content; }
}
$e = new Editor();
$e->content = 'version1';
$snap = $e->save();
$e->content = 'version2';
echo $e->content;
$e->restore($snap);
echo $e->content;

__vybe_check(ob_get_clean(), "version2version1");
