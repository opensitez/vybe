<?php
// vybe-test: php/namespaces/use_group
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace Domain\Models;
class Order   { public string $id = 'ORD'; }
class Product { public string $id = 'PRD'; }
class Invoice { public string $id = 'INV'; }

namespace App;
use Domain\Models\{Order, Product, Invoice};
$o = new Order();
$p = new Product();
$i = new Invoice();
echo $o->id . $p->id . $i->id;
