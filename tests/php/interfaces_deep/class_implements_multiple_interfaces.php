<?php
// vybe-test: php/interfaces_deep/class_implements_multiple_interfaces
// origin: languages/php/tests/php/test_interfaces_deep.rs
// vybe-test-mode: compile

interface Printable  { public function print(): void; }
interface Saveable   { public function save(): bool; }
interface Deletable  { public function delete(): bool; }
class Record implements Printable, Saveable, Deletable {
    public function print(): void  { echo 'printing'; }
    public function save(): bool   { return true; }
    public function delete(): bool { return true; }
}
$r = new Record();
$r->print();
echo $r->save()   ? ':saved'   : ':save failed';
echo $r->delete() ? ':deleted' : ':delete failed';
