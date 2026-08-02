<?php
// vybe-test: php/named_arguments/named_args_with_trait_method
// origin: languages/php/tests/php/test_named_arguments.rs

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

trait Formattable {
    public function format(string $template, string $locale = 'en'): string {
        return str_replace('{locale}', $locale, str_replace('{val}', (string)$this->value, $template));
    }
}
class Temperature {
    use Formattable;
    public function __construct(public float $value) {}
}
$t = new Temperature(value: 23.5);
echo $t->format(template: '{val} ({locale})', locale: 'en') . "\n";

__vybe_check(ob_get_clean(), "23.5 (en)");
