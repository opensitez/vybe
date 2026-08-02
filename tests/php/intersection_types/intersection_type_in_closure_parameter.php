<?php
// vybe-test: php/intersection_types/intersection_type_in_closure_parameter
// origin: languages/php/tests/php/test_intersection_types.rs
// vybe-test-mode: compile

interface Readable2 { public function read(): string; }
interface Seekable { public function seek(int $pos): void; }
$process = function(Readable2&Seekable $stream): string {
    $stream->seek(0);
    return $stream->read();
};
