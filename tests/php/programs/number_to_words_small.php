<?php
// vybe-test: php/programs/number_to_words_small
// origin: languages/php/tests/php/test_programs.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

function numberToWords(int $n): string {
    $ones = ['','one','two','three','four','five','six','seven','eight','nine',
             'ten','eleven','twelve','thirteen','fourteen','fifteen','sixteen',
             'seventeen','eighteen','nineteen'];
    $tens = ['','','twenty','thirty','forty','fifty','sixty','seventy','eighty','ninety'];
    if ($n < 20) return $ones[$n];
    if ($n < 100) return $tens[intdiv($n,10)] . ($n%10 ? '-' . $ones[$n%10] : '');
    return $ones[intdiv($n,100)] . ' hundred' . ($n%100 ? ' ' . numberToWords($n%100) : '');
}
echo numberToWords(42) . "\n";
echo numberToWords(7) . "\n";
echo numberToWords(100) . "\n";

__vybe_check(ob_get_clean(), "forty-two\nseven\none hundred");
