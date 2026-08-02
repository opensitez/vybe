<?php
// vybe-test: php/spl_extra/spl_fixed_array_set_size
// origin: languages/php/tests/php/test_spl_extra.rs
// vybe-test-mode: compile

$fa = new SplFixedArray(3);
$fa[0] = 'x'; $fa[1] = 'y'; $fa[2] = 'z';
$fa->setSize(5);
$fa[3] = 'a'; $fa[4] = 'b';
echo $fa->getSize();
echo $fa[4];
