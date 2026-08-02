<?php
// vybe-test: php/design_patterns/command_pattern_undo
// origin: languages/php/tests/php/test_design_patterns.rs

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

interface Command { public function execute(): void; public function undo(): void; }
class Stack {
    private array $stack = [];
    private array $history = [];
    public function execute(Command $cmd): void { $cmd->execute(); $this->history[] = $cmd; }
    public function undo(): void { if ($c = array_pop($this->history)) $c->undo(); }
}
class PushCommand implements Command {
    public function __construct(private array &$list, private int $val) {}
    public function execute(): void { $this->list[] = $this->val; }
    public function undo(): void { array_pop($this->list); }
}
$list = [];
$stack = new Stack;
$stack->execute(new PushCommand($list, 1));
$stack->execute(new PushCommand($list, 2));
$stack->execute(new PushCommand($list, 3));
echo implode(',', $list) . ',';
$stack->undo();
echo implode(',', $list);

__vybe_check(ob_get_clean(), "1,2,3,1,2");
