<?php
// vybe-test: php/spl_extra/spl_fixed_array_from_regular_array
// origin: languages/php/tests/php/test_spl_extra.rs
// vybe-test-mode: compile

$fa = SplFixedArray::fromArray([100, 200, 300, 400]);
echo $fa->getSize();
echo $fa[3];
