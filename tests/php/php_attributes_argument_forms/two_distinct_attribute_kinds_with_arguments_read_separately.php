<?php
// vybe-test: php/php_attributes_argument_forms/two_distinct_attribute_kinds_with_arguments_read_separately
// origin: languages/php/tests/php/test_php_attributes_argument_forms.rs

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
class Route {
    public function __construct(public string $path) {}
}
#[Attribute]
class Auth {
    public function __construct(public string $role) {}
}
class Admin {
    #[Route('/admin')]
    #[Auth('superuser')]
    public function panel() {}
}
$rm = new ReflectionMethod(Admin::class, 'panel');
echo $rm->getAttributes(Route::class)[0]->newInstance()->path
   . '+' . $rm->getAttributes(Auth::class)[0]->newInstance()->role;

__vybe_check(ob_get_clean(), "/admin+superuser");
