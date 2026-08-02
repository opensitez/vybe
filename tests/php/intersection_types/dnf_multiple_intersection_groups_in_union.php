<?php
// vybe-test: php/intersection_types/dnf_multiple_intersection_groups_in_union
// origin: languages/php/tests/php/test_intersection_types.rs
// vybe-test-mode: compile

interface A2 { public function a(): void; }
interface B2 { public function b(): void; }
interface C2 { public function c(): void; }
interface D2 { public function d(): void; }
function process((A2&B2)|(C2&D2) $obj): void {}
