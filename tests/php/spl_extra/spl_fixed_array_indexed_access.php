<?php
// vybe-test: php/spl_extra/spl_fixed_array_indexed_access
// origin: languages/php/tests/php/test_spl_extra.rs
// vybe-test-mode: compile

$fa = new SplFixedArray(4);
$fa[0] = 10; $fa[1] = 20; $fa[2] = 30; $fa[3] = 40;
echo $fa->getSize() . ':' . $fa[2];
