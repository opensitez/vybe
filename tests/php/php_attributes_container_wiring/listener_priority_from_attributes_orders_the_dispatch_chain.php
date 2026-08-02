<?php
// vybe-test: php/php_attributes_container_wiring/listener_priority_from_attributes_orders_the_dispatch_chain
// origin: languages/php/tests/php/test_php_attributes_container_wiring.rs

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
class AsListener {
    public function __construct(public string $event, public int $priority = 0) {}
}
#[AsListener('kernel.request', priority: 10)]
class Auth {}
#[AsListener('kernel.request', priority: 100)]
class Firewall {}
#[AsListener('kernel.request')]
class Router {}
$ls = [];
foreach ([Auth::class, Firewall::class, Router::class] as $c) {
    $a = (new ReflectionClass($c))->getAttributes(AsListener::class)[0]->newInstance();
    $ls[] = [$c, $a->priority];
}
usort($ls, fn($x, $y) => $y[1] <=> $x[1]);
echo implode('>', array_map(fn($l) => $l[0] . '(' . $l[1] . ')', $ls));

__vybe_check(ob_get_clean(), "Firewall(100)>Auth(10)>Router(0)");
