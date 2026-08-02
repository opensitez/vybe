<?php
// vybe-test: php/php_attributes_inheritance_traits/trait_method_attributes_are_visible_on_the_using_class
// origin: languages/php/tests/php/test_php_attributes_inheritance_traits.rs

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
class Hook {
    public function __construct(public string $when) {}
}
trait Boots {
    #[Hook('boot')]
    public function boot() {}
}
class App {
    use Boots;
}
echo (new ReflectionMethod(App::class, 'boot'))->getAttributes(Hook::class)[0]->newInstance()->when;

__vybe_check(ob_get_clean(), "boot");
