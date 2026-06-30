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

// ── Runtime namespace resolution (`php_cases!`) ─────────────────

crate::php_cases! {
    namespace_class_instantiated_with_leading_backslash => {
        r#"<?php
namespace App\Models {
    class User {
        public function __construct(public string $name) {}
    }
}
echo (new \App\Models\User('ada'))->name;
"#,
        ["ada"]
    };

    namespace_function_called_from_global => {
        r#"<?php
namespace App\Support {
    function greet(string $who): string {
        return 'hi:' . $who;
    }
}
echo \App\Support\greet('vybe');
"#,
        ["hi:vybe"]
    };

    namespace_constant_read_from_global => {
        r#"<?php
namespace Config {
    const APP_NAME = 'VybeApp';
}
echo \Config\APP_NAME;
"#,
        ["VybeApp"]
    };

    use_import_shortens_class_reference => {
        r#"<?php
namespace App\Http {
    class Request {}
}
namespace App\Controllers {
    use App\Http\Request;
    function make(): string {
        return (new Request()) instanceof Request ? 'req' : 'no';
    }
}
echo \App\Controllers\make();
"#,
        ["req"]
    };

    use_function_imports_namespaced_function => {
        r#"<?php
namespace Lib {
    function twice(int $n): int { return $n * 2; }
}
namespace App {
    use function Lib\twice;
    echo twice(4);
}
"#,
        ["8"]
    };

    use_const_imports_namespaced_constant => {
        r#"<?php
namespace Lib {
    const MAX = 100;
}
namespace App {
    use const Lib\MAX;
    echo MAX;
}
"#,
        ["100"]
    };

    nested_namespace_segments => {
        r#"<?php
namespace Vendor\Package\Sub {
    class Tool {
        public function tag(): string { return 'tool'; }
    }
}
echo (new \Vendor\Package\Sub\Tool())->tag();
"#,
        ["tool"]
    };

    class_from_same_namespace_without_prefix => {
        r#"<?php
namespace App {
    class A { public function v(): int { return 1; } }
    class B {
        public function pull(): int {
            return (new A())->v();
        }
    }
}
echo (new \App\B())->pull();
"#,
        ["1"]
    };

    static_call_within_namespace => {
        r#"<?php
namespace Math {
    class Calc {
        public static function add(int $a, int $b): int { return $a + $b; }
    }
}
echo \Math\Calc::add(2, 3);
"#,
        ["5"]
    };

    enum_in_namespace => {
        r#"<?php
namespace App\Enums {
    enum Status { case Active; case Draft; }
}
echo \App\Enums\Status::Draft->name;
"#,
        ["Draft"]
    };

    trait_used_inside_namespaced_class => {
        r#"<?php
namespace App\Traits {
    trait Timestamped {
        public function stamp(): string { return 'ts'; }
    }
}
namespace App\Models {
    use App\Traits\Timestamped;
    class Post {
        use Timestamped;
    }
}
echo (new \App\Models\Post())->stamp();
"#,
        ["ts"]
    };

    interface_implementation_across_namespaces => {
        r#"<?php
namespace Contracts {
    interface Repository { public function all(): array; }
}
namespace Infra {
    class MemoryRepo implements \Contracts\Repository {
        public function all(): array { return [1, 2]; }
    }
}
echo count((new \Infra\MemoryRepo())->all());
"#,
        ["2"]
    };

    fully_qualified_name_bypasses_use_conflict => {
        r#"<?php
namespace A { class Name { public function id(): string { return 'A'; } } }
namespace B { class Name { public function id(): string { return 'B'; } } }
namespace App {
    use A\Name;
    function pick(): string {
        $a = new Name();
        $b = new \B\Name();
        return $a->id() . $b->id();
    }
}
echo \App\pick();
"#,
        ["AB"]
    };

    namespace_group_use_braces => {
        r#"<?php
namespace Vendor\Support {
    class Str {
        public static function upper(string $s): string { return strtoupper($s); }
    }
}
namespace App {
    use Vendor\Support\{Str};
    echo Str::upper('ok');
}
"#,
        ["OK"]
    };

    global_class_referenced_from_namespace => {
        r#"<?php
namespace App {
    function make(): string {
        $o = new \stdClass();
        $o->x = 'std';
        return $o->x;
    }
}
echo \App\make();
"#,
        ["std"]
    };

    namespace_same_name_different_segments => {
        r#"<?php
namespace Foo\Bar { class X { public function t(): string { return 'fb'; } } }
namespace Foo\Baz { class X { public function t(): string { return 'fz'; } } }
echo (new \Foo\Bar\X())->t() . (new \Foo\Baz\X())->t();
"#,
        ["fbfz"]
    };

    namespace_magic_constant_echoes_fully_qualified_name => {
        r#"<?php
namespace Vendor\Package {
    function ns(): string { return __NAMESPACE__; }
}
echo \Vendor\Package\ns();
"#,
        ["Vendor\\Package"]
    };

    use_statement_alias_shortens_runtime_reference => {
        r#"<?php
namespace Lib\Collections {
    class Bag {
        public function size(): int { return 2; }
    }
}
namespace App {
    use Lib\Collections\Bag as Container;
    function count_items(): int {
        return (new Container())->size();
    }
}
echo \App\count_items();
"#,
        ["2"]
    };

    multiple_namespace_blocks_in_one_file => {
        r#"<?php
namespace Alpha { class Node { public function tag(): string { return 'A'; } } }
namespace Beta { class Node { public function tag(): string { return 'B'; } } }
echo (new \Alpha\Node())->tag() . (new \Beta\Node())->tag();
"#,
        ["AB"]
    };

    braced_namespace_with_inner_declarations => {
        r#"<?php
namespace App\Models {
    class User {
        public function __construct(public string $name) {}
    }
}
echo (new \App\Models\User('leo'))->name;
"#,
        ["leo"]
    };

    global_namespace_block_accesses_both_namespaces => {
        r#"<?php
namespace N1 { class S { public function v(): int { return 1; } } }
namespace N2 { class S { public function v(): int { return 2; } } }
namespace {
    function sum(): int {
        return (new \N1\S())->v() + (new \N2\S())->v();
    }
}
echo sum();
"#,
        ["3"]
    };

    use_group_imports_multiple_classes_runtime => {
        r#"<?php
namespace Parts {
    class Wheel { public function id(): string { return 'W'; } }
    class Axle { public function id(): string { return 'A'; } }
}
namespace Garage {
    use Parts\{Wheel, Axle};
    function ids(): string {
        return (new Wheel())->id() . (new Axle())->id();
    }
}
echo \Garage\ids();
"#,
        ["WA"]
    };

    use_group_mixed_imports_function_and_const => {
        r#"<?php
namespace Util {
    function wrap(string $s): string { return "[$s]"; }
    const LEVEL = 'debug';
}
namespace App {
    use Util\{function wrap, const LEVEL};
    function line(): string { return wrap(LEVEL); }
}
echo \App\line();
"#,
        ["[debug]"]
    };

    parent_namespace_relative_class_not_used_fqcn_wins => {
        r#"<?php
namespace Project\Core {
    class Engine { public function rev(): string { return 'v8'; } }
}
namespace Project\App {
    class Car {
        public function engine(): string {
            return (new \Project\Core\Engine())->rev();
        }
    }
}
echo (new \Project\App\Car())->engine();
"#,
        ["v8"]
    };

    namespace_function_sees_own_namespace_for_unqualified_class => {
        r#"<?php
namespace Shop {
    class Item { public function sku(): string { return 'SKU'; } }
    function make(): string { return (new Item())->sku(); }
}
echo \Shop\make();
"#,
        ["SKU"]
    };

    leading_backslash_builtin_from_inside_namespace => {
        r#"<?php
namespace App {
    function len(string $s): int { return \strlen($s); }
}
echo \App\len('php');
"#,
        ["3"]
    };

    namespace_collision_resolved_by_fully_qualified_import => {
        r#"<?php
namespace Legacy { class Logger { public function id(): string { return 'L'; } } }
namespace Modern { class Logger { public function id(): string { return 'M'; } } }
namespace App {
    use Legacy\Logger;
    function pick(): string {
        $old = new Logger();
        $new = new \Modern\Logger();
        return $old->id() . $new->id();
    }
}
echo \App\pick();
"#,
        ["LM"]
    };
}
