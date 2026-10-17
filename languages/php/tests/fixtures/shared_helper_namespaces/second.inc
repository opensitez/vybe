<?php
namespace HelperSecond;

class Same {
    public function value() { return 'second'; }
}

function label($value) {
    return $value == 7 ? 'second-label' : 'wrong';
}
