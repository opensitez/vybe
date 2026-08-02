<?php
// vybe-test: php/namespaces/multiple_namespaces_one_file
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace Alpha;
class Foo { public function name(): string { return 'Foo'; } }

namespace Beta;
class Bar { public function name(): string { return 'Bar'; } }

namespace {
    $a = new \Alpha\Foo();
    $b = new \Beta\Bar();
    echo $a->name() . $b->name();
}
