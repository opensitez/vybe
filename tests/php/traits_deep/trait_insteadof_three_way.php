<?php
// vybe-test: php/traits_deep/trait_insteadof_three_way
// origin: languages/php/tests/php/test_traits_deep.rs
// vybe-test-mode: compile

trait X { public function op(): string { return "X"; } }
trait Y { public function op(): string { return "Y"; } }
trait Z { public function op(): string { return "Z"; } }
class UseXY {
    use X, Y, Z { X::op insteadof Y, Z; Y::op as opY; Z::op as opZ; }
}
$u = new UseXY();
echo $u->op() . $u->opY() . $u->opZ();
