<?php
// vybe-test: php/php_attributes_container_wiring/constructor_parameter_attributes_drive_container_resolution
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

#[Attribute(Attribute::TARGET_PARAMETER)]
class Autowire {
    public function __construct(public string $service) {}
}
class Mailer {
    public function send(): string { return 'sent'; }
}
class Notifier {
    public function __construct(
        #[Autowire('mailer')] public Mailer $m,
        #[Autowire('logger')] public $log
    ) {}
}
$container = ['mailer' => new Mailer(), 'logger' => 'LOG'];
$args = [];
foreach ((new ReflectionClass(Notifier::class))->getConstructor()->getParameters() as $p) {
    $a = $p->getAttributes(Autowire::class);
    $args[] = $a ? $container[$a[0]->newInstance()->service] : null;
}
$n = new Notifier(...$args);
echo $n->m->send() . '/' . $n->log;

__vybe_check(ob_get_clean(), "sent/LOG");
