<?php
// vybe-test: php/interfaces_deep/polymorphism_collection
// origin: languages/php/tests/php/test_interfaces_deep.rs
// vybe-test-mode: compile

interface Formatter { public function format(mixed $value): string; }
class IntFormatter implements Formatter {
    public function format(mixed $value): string { return number_format((int)$value); }
}
class DateFormatter implements Formatter {
    public function format(mixed $value): string { return date('Y-m-d', (int)$value); }
}
class BoolFormatter implements Formatter {
    public function format(mixed $value): string { return $value ? 'yes' : 'no'; }
}
/** @param Formatter[] $formatters */
function formatAll(array $data, array $formatters): array {
    $result = [];
    foreach ($data as $i => $v) {
        $result[] = ($formatters[$i] ?? $formatters[0])->format($v);
    }
    return $result;
}
$rows = formatAll([1000, true], [new IntFormatter(), new BoolFormatter()]);
echo implode('|', $rows);
