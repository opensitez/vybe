<?php
// vybe-test: php/iterators/stringable_basic
// origin: languages/php/tests/php/test_iterators.rs
// vybe-test-mode: compile

class Version implements Stringable {
    public function __construct(
        private int $major,
        private int $minor,
        private int $patch
    ) {}
    public function __toString(): string { return "{$this->major}.{$this->minor}.{$this->patch}"; }
}
function printVersion(Stringable $v): void { echo (string)$v; }
printVersion(new Version(1, 2, 3));
