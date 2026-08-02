<?php
// vybe-test: php/patterns/interceptor_wrap_method_calls
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

class ServiceProxy {
    private $service;
    private $callLog = [];
    public function __construct(object $service) { $this->service = $service; }
    public function __call(string $method, array $args) {
        $this->callLog[] = $method;
        echo 'before:' . $method;
        $result = $this->service->$method(...$args);
        echo 'after:' . $method;
        return $result;
    }
    public function getCallLog(): array { return $this->callLog; }
}
class RealService {
    public function doWork(string $task): string { echo 'working:' . $task; return 'done'; }
}
$proxy = new ServiceProxy(new RealService());
$proxy->doWork('task1');
echo implode(',', $proxy->getCallLog());

__vybe_check(ob_get_clean(), "before:doWorkworking:task1after:doWorkdoWork");
