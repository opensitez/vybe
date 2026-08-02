<?php
// vybe-test: php/php_traits_composition_conflict_resolution/test_php_trait_aliasing_retaining_original_name
// origin: languages/php/tests/php/test_php_traits_composition_conflict_resolution.rs
// vybe-test-mode: compile

trait OutputHelper {
    public function print() { return "Original Print"; }
}

class Printer {
    use OutputHelper {
        print as customPrint;
    }
}

$p = new Printer();
echo $p->print() . " | " . $p->customPrint();
