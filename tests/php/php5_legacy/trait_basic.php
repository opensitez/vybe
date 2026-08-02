<?php
// vybe-test: php/php5_legacy/trait_basic
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

trait Timestampable {
    public function getCreated() { return $this->created; }
}
class Post { use Timestampable; public $created = '2024-01-01'; }
$p = new Post();
echo $p->getCreated();
