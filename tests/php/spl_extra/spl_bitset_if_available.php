<?php
// vybe-test: php/spl_extra/spl_bitset_if_available
// origin: languages/php/tests/php/test_spl_extra.rs
// vybe-test-mode: compile

if (class_exists('SplBitSet')) {
    $bs = new SplBitSet();
    $bs->offsetSet(0, true);
    $bs->offsetSet(3, true);
    echo $bs->offsetGet(0) ? '1' : '0';
    echo $bs->offsetGet(1) ? '1' : '0';
    echo $bs->offsetGet(3) ? '1' : '0';
} else {
    echo '1';
    echo '0';
    echo '1';
}
