<?php
// vybe-test: php/php_array_sorting_multisort_callbacks/test_php_natcasesort_natural_order_sorting
// origin: languages/php/tests/php/test_php_array_sorting_multisort_callbacks.rs
// vybe-test-mode: compile

$files = ["img12.png", "img10.png", "img2.png", "img1.png"];
natcasesort($files);
echo implode(",", $files);
