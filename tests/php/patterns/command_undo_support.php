<?php
// vybe-test: php/patterns/command_undo_support
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

class TextEditor {
    public $text = '';
    public function append(string $s): void { $this->text .= $s; }
    public function deleteLast(int $n): void { $this->text = substr($this->text, 0, strlen($this->text) - $n); }
}
interface Command {
    public function execute(): void;
    public function undo(): void;
}
class AppendCommand implements Command {
    private $editor;
    private $text;
    public function __construct(TextEditor $e, string $t) { $this->editor = $e; $this->text = $t; }
    public function execute(): void { $this->editor->append($this->text); }
    public function undo(): void { $this->editor->deleteLast(strlen($this->text)); }
}
$editor = new TextEditor();
$history = [];
$c1 = new AppendCommand($editor, 'Hello');
$c1->execute();
$history[] = $c1;
$c2 = new AppendCommand($editor, ' World');
$c2->execute();
$history[] = $c2;
echo $editor->text;
array_pop($history)->undo();
echo $editor->text;

__vybe_check(ob_get_clean(), "Hello WorldHello");
