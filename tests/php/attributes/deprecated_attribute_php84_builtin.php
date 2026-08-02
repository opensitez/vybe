<?php
// vybe-test: php/attributes/deprecated_attribute_php84_builtin
// origin: languages/php/tests/php/test_attributes.rs

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

class Legacy {
    #[\Deprecated("Use newMethod instead", since: "2.0")]
    public function oldMethod(): string { return 'legacy'; }
}
$rm = new ReflectionMethod(Legacy::class, 'oldMethod');
$attrs = $rm->getAttributes(\Deprecated::class);
echo count($attrs) . '|' . $attrs[0]->getName() . '|';
$d = $attrs[0]->newInstance();
echo $d->message . '|' . $d->since . '|' . ($rm->isDeprecated() ? 'yes' : 'no');

__vybe_check(ob_get_clean(), "1|Deprecated|Use newMethod instead|2.0|yes");
