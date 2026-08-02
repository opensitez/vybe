<?php
// vybe-test: php/bcmath/bcmath_factorial_large
// origin: languages/php/tests/php/test_bcmath.rs
// vybe-test-mode: compile

function bcfactorial(int $n): string {
    $result = '1';
    for ($i = 2; $i <= $n; $i++) {
        $result = bcmul($result, (string)$i);
    }
    return $result;
}
$f20 = bcfactorial(20);
echo strlen($f20) > 10 ? 'large factorial' : 'too small';
echo bccomp($f20, '2432902008176640000') >= 0 ? ':correct magnitude' : ':wrong';
