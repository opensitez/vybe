<?php
// vybe-test: php/enums_deep/enum_method_comparison
// origin: languages/php/tests/php/test_enums_deep.rs
// vybe-test-mode: compile

enum Size: int {
    case XS = 1; case S = 2; case M = 3; case L = 4; case XL = 5;
    public function fitsInto(self $other): bool { return $this->value <= $other->value; }
    public function between(self $min, self $max): bool {
        return $this->value >= $min->value && $this->value <= $max->value;
    }
}
echo Size::S->fitsInto(Size::L) ? 'fits' : 'no fit';
echo Size::M->between(Size::S, Size::XL) ? ':in range' : ':out of range';
