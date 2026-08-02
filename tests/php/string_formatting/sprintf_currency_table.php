<?php
// vybe-test: php/string_formatting/sprintf_currency_table
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

$items = [['Widget', 5, 9.99], ['Gadget', 2, 24.95], ['Doohickey', 1, 4.50]];
$total = 0.0;
foreach ($items as [$name, $qty, $price]) {
    $line = $qty * $price;
    $total += $line;
    printf("%-12s %3d @ %6.2f = %8.2f\n", $name, $qty, $price, $line);
}
printf("%s\n", str_repeat('-', 36));
printf("%-18s %16.2f\n", 'Total:', $total);
