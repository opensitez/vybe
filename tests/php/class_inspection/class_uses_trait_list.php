<?php
// vybe-test: php/class_inspection/class_uses_trait_list
// origin: languages/php/tests/php/test_class_inspection.rs
// vybe-test-mode: compile

trait HasUuid { public function uuid(): string { return 'abc'; } }
trait HasTimestamp { public function ts(): int { return 0; } }
class Record {
    use HasUuid, HasTimestamp;
}
$traits = class_uses('Record');
echo isset($traits['HasUuid']) ? 'yes' : 'no';
echo isset($traits['HasTimestamp']) ? 'yes' : 'no';
