<?php
// vybe-test: php/host_extra/hex_color
// origin: languages/php/tests/php/test_host_extra.rs
// vybe-test-mode: compile

function hexColor($r, $g, $b) {
    return '#' . str_pad(dechex($r), 2, '0') . str_pad(dechex($g), 2, '0') . str_pad(dechex($b), 2, '0');
}
echo hexColor(255, 128, 0);
