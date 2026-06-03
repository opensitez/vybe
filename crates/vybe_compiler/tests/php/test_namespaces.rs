use super::helpers::compile_ok;

// ── Basic namespace declarations ──────────────────────────────

#[test]
fn namespace_basic() {
    compile_ok(
        r#"<?php
namespace App;
class Greeter {
    public function greet(string $name): string {
        return "Hello, $name!";
    }
}
$g = new Greeter();
echo $g->greet("World");
"#,
    );
}

#[test]
fn namespace_function() {
    compile_ok(
        r#"<?php
namespace Utils;
function clamp(int $v, int $lo, int $hi): int {
    return max($lo, min($hi, $v));
}
echo clamp(15, 0, 10);
"#,
    );
}

#[test]
fn namespace_constant() {
    compile_ok(
        r#"<?php
namespace Config;
const VERSION = '1.0.0';
const MAX_RETRIES = 3;
echo VERSION . ' retries=' . MAX_RETRIES;
"#,
    );
}

#[test]
fn namespace_nested() {
    compile_ok(
        r#"<?php
namespace App\Http\Request;
class Handler {
    public function handle(string $method): string {
        return strtoupper($method);
    }
}
$h = new Handler();
echo $h->handle('get');
"#,
    );
}

// ── use statement ─────────────────────────────────────────────

#[test]
fn use_class() {
    compile_ok(
        r#"<?php
namespace Models;
class User {
    public function __construct(public string $name) {}
}

namespace App;
use Models\User;
$u = new User('Alice');
echo $u->name;
"#,
    );
}

#[test]
fn use_with_alias() {
    compile_ok(
        r#"<?php
namespace Library;
class Collection {
    private array $items = [];
    public function add(mixed $v): void { $this->items[] = $v; }
    public function count(): int { return count($this->items); }
}

namespace App;
use Library\Collection as List_;
$list = new List_();
$list->add(1); $list->add(2);
echo $list->count();
"#,
    );
}

#[test]
fn use_function() {
    compile_ok(
        r#"<?php
namespace Helpers;
function slugify(string $s): string {
    return strtolower(str_replace(' ', '-', $s));
}

namespace App;
use function Helpers\slugify;
echo slugify('Hello World');
"#,
    );
}

#[test]
fn use_const() {
    compile_ok(
        r#"<?php
namespace Constants;
const PI = 3.14159;
const E  = 2.71828;

namespace App;
use const Constants\PI;
use const Constants\E;
echo round(PI, 2) . ',' . round(E, 2);
"#,
    );
}

#[test]
fn use_group() {
    compile_ok(
        r#"<?php
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
"#,
    );
}

#[test]
fn use_group_mixed() {
    compile_ok(
        r#"<?php
namespace Util;
class Logger { public function log(string $m): void { echo $m; } }
function format(string $s): string { return "[$s]"; }
const LOG_LEVEL = 'info';

namespace App;
use Util\{Logger, function format, const LOG_LEVEL};
$log = new Logger();
$log->log(format(LOG_LEVEL));
"#,
    );
}

// ── __NAMESPACE__ constant ────────────────────────────────────

#[test]
fn namespace_magic_constant() {
    compile_ok(
        r#"<?php
namespace Vendor\Package;
echo __NAMESPACE__;
"#,
    );
}

#[test]
fn namespace_magic_in_function() {
    compile_ok(
        r#"<?php
namespace App\Core;
function getNamespace(): string { return __NAMESPACE__; }
echo getNamespace();
"#,
    );
}

// ── Global namespace escape ───────────────────────────────────

#[test]
fn global_namespace_backslash() {
    compile_ok(
        r#"<?php
namespace App;
$len = \strlen("hello");
echo $len;
"#,
    );
}

#[test]
fn global_namespace_class() {
    compile_ok(
        r#"<?php
namespace App;
$e = new \Exception("from global ns");
echo $e->getMessage();
"#,
    );
}

// ── Namespace in OOP ──────────────────────────────────────────

#[test]
fn namespace_interface() {
    compile_ok(
        r#"<?php
namespace Contracts;
interface Renderable {
    public function render(): string;
}

namespace UI;
use Contracts\Renderable;
class Button implements Renderable {
    public function __construct(private string $label) {}
    public function render(): string { return "<button>{$this->label}</button>"; }
}
$btn = new Button("Click me");
echo $btn->render();
"#,
    );
}

#[test]
fn namespace_trait() {
    compile_ok(
        r#"<?php
namespace Concerns;
trait Timestampable {
    private int $createdAt = 0;
    public function setCreatedAt(int $ts): void { $this->createdAt = $ts; }
    public function getCreatedAt(): int { return $this->createdAt; }
}

namespace Models;
use Concerns\Timestampable;
class Post {
    use Timestampable;
    public function __construct(public string $title) {}
}
$p = new Post('Hello');
$p->setCreatedAt(1000);
echo $p->getCreatedAt();
"#,
    );
}

#[test]
fn namespace_enum() {
    compile_ok(
        r#"<?php
namespace Domain\Status;
enum OrderStatus: string {
    case Pending  = 'pending';
    case Shipped  = 'shipped';
    case Delivered = 'delivered';
}

namespace App;
use Domain\Status\OrderStatus;
$status = OrderStatus::Shipped;
echo $status->value;
"#,
    );
}

#[test]
fn namespace_abstract_class() {
    compile_ok(
        r#"<?php
namespace Base;
abstract class Shape {
    abstract public function area(): float;
    public function describe(): string { return "area=" . $this->area(); }
}

namespace Shapes;
use Base\Shape;
class Circle extends Shape {
    public function __construct(private float $r) {}
    public function area(): float { return M_PI * $this->r ** 2; }
}
$c = new Circle(2.0);
echo round($c->area(), 4);
"#,
    );
}

// ── Multiple namespaces in one file ───────────────────────────

#[test]
fn multiple_namespaces_one_file() {
    compile_ok(
        r#"<?php
namespace Alpha;
class Foo { public function name(): string { return 'Foo'; } }

namespace Beta;
class Bar { public function name(): string { return 'Bar'; } }

namespace {
    $a = new \Alpha\Foo();
    $b = new \Beta\Bar();
    echo $a->name() . $b->name();
}
"#,
    );
}

// ── Namespace collision handling ──────────────────────────────

#[test]
fn namespace_same_class_name_different_ns() {
    compile_ok(
        r#"<?php
namespace V1;
class Response { public function status(): int { return 200; } }

namespace V2;
class Response { public function status(): int { return 201; } }

namespace App;
$r1 = new \V1\Response();
$r2 = new \V2\Response();
echo $r1->status() . ',' . $r2->status();
"#,
    );
}

#[test]
fn namespace_fully_qualified_call() {
    compile_ok(
        r#"<?php
namespace Helpers;
function double(int $n): int { return $n * 2; }

namespace App;
$result = \Helpers\double(21);
echo $result;
"#,
    );
}
