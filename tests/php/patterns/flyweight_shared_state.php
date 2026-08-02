<?php
// vybe-test: php/patterns/flyweight_shared_state
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

class TreeType {
    public function __construct(public string $name, public string $color) {}
    public function draw(int $x, int $y): string { return "{$this->name}@{$x},{$y}"; }
}
class TreeFactory {
    private static $types = [];
    public static function get(string $name, string $color): TreeType {
        $key = "$name-$color";
        if (!isset(self::$types[$key])) {
            self::$types[$key] = new TreeType($name, $color);
        }
        return self::$types[$key];
    }
    public static function count(): int { return count(self::$types); }
}
$t1 = TreeFactory::get('oak', 'green');
$t2 = TreeFactory::get('oak', 'green');
$t3 = TreeFactory::get('pine', 'dark-green');
echo ($t1 === $t2) ? 'shared' : 'different';
echo TreeFactory::count();
echo $t1->draw(1, 2);

__vybe_check(ob_get_clean(), "shared2oak@1,2");
