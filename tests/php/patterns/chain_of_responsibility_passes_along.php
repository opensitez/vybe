<?php
// vybe-test: php/patterns/chain_of_responsibility_passes_along
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

abstract class Handler {
    private $next = null;
    public function setNext(Handler $h): Handler { $this->next = $h; return $h; }
    public function handle(int $req): string {
        if ($this->next !== null) return $this->next->handle($req);
        return 'unhandled';
    }
}
class SmallHandler extends Handler {
    public function handle(int $req): string {
        if ($req < 10) return 'small:' . $req;
        return parent::handle($req);
    }
}
class MediumHandler extends Handler {
    public function handle(int $req): string {
        if ($req < 100) return 'medium:' . $req;
        return parent::handle($req);
    }
}
class LargeHandler extends Handler {
    public function handle(int $req): string { return 'large:' . $req; }
}
$small = new SmallHandler();
$medium = new MediumHandler();
$large = new LargeHandler();
$small->setNext($medium)->setNext($large);
echo $small->handle(5);
echo $small->handle(50);
echo $small->handle(500);

__vybe_check(ob_get_clean(), "small:5medium:50large:500");
