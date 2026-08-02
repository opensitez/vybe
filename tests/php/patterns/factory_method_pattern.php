<?php
// vybe-test: php/patterns/factory_method_pattern
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

abstract class Transport {
    abstract public function getType(): string;
}
class Truck extends Transport {
    public function getType(): string { return 'truck'; }
}
class Ship extends Transport {
    public function getType(): string { return 'ship'; }
}
abstract class Logistics {
    abstract public function createTransport(): Transport;
    public function plan(): string { return $this->createTransport()->getType(); }
}
class RoadLogistics extends Logistics {
    public function createTransport(): Transport { return new Truck(); }
}
class SeaLogistics extends Logistics {
    public function createTransport(): Transport { return new Ship(); }
}
echo (new RoadLogistics())->plan();
echo (new SeaLogistics())->plan();

__vybe_check(ob_get_clean(), "truckship");
