<?php
namespace HelperFirst;

class Same {
    public function value() { return 'first'; }
}

function label($value) {
    return $value == 7 ? 'first-label' : 'wrong';
}
