<?php
// vybe-test: php/magic_constants/magic_all_in_class
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

class Diagnostics {
    public function report(): array {
        return [
            'class'    => __CLASS__,
            'method'   => __METHOD__,
            'line'     => __LINE__,
            'file_set' => __FILE__ !== '',
        ];
    }
}
$info = (new Diagnostics())->report();
echo $info['class'] . ':' . $info['method'];
echo ':line=' . ($info['line'] > 0 ? 'ok' : 'fail');
echo ':file=' . ($info['file_set'] ? 'ok' : 'fail');
