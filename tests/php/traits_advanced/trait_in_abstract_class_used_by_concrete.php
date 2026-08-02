<?php
// vybe-test: php/traits_advanced/trait_in_abstract_class_used_by_concrete
// origin: languages/php/tests/php/test_traits_advanced.rs

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

trait EventEmitter {
    private array $listeners = [];
    public function on(string $event, callable $cb): void { $this->listeners[$event][] = $cb; }
    public function emit(string $event, mixed ...$args): void {
        foreach ($this->listeners[$event] ?? [] as $cb) $cb(...$args);
    }
}
abstract class Component { use EventEmitter; abstract public function render(): string; }
class Button extends Component {
    public function render(): string { return '<button>'; }
}
$btn = new Button;
$btn->on('click', fn($x) => print("clicked:$x"));
$btn->emit('click', 'left');

__vybe_check(ob_get_clean(), "clicked:left");
