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

    namespace_rootless_class_lookup_in_parent_segment => {
        r#"<?php
namespace App\Core {
    class Engine {
        public function version(): string { return 'v1'; }
    }
}
namespace App {
    function getEngineVersion(): string {
        $e = new Core\Engine();
        return $e->version();
    }
}
echo \App\getEngineVersion();
"#,
        ["v1"]
    };

    namespace_current_namespace_prefers_local_function => {
        r#"<?php
namespace Lib {
    function format(string $s): string { return 'local:' . $s; }
}
namespace App {
    function format(string $s): string { return 'app:' . $s; }
    function run(): string {
        return \Lib\format('a') . '|' . format('b');
    }
}
echo \App\run();
"#,
        ["local:a|app:b"]
    };

    namespace_fully_qualified_trait_use => {
        r#"<?php
namespace Framework\Traits {
    trait Timestamped {
        public function timestamp(): string { return 'ts'; }
    }
}
namespace App {
    class Entity {
        use \Framework\Traits\Timestamped;
    }
}
echo (new \App\Entity())->timestamp();
"#,
        ["ts"]
    };

    namespace_aliasing_function_chain => {
        r#"<?php
namespace Math {
    function add(int $a, int $b): int { return $a + $b; }
}
namespace App {
    use Math\add as plus;
    echo plus(2, 5);
}
"#,
        ["7"]
    };

    namespace_dynamic_fully_qualified_function_call => {
        r#"<?php
namespace Runtime {
    function toString(int $v): string { return 'v=' . $v; }
}
namespace App {
    $fqn = 'Runtime\\toString';
    $name = '\\\\' . $fqn;
    echo $name(9);
}
"#,
        ["v=9"]
    };

    namespace_relative_vs_global_function_resolution => {
        r#"<?php
function marker(string $s): string { return 'global:' . $s; }
namespace App {
    function marker(string $s): string { return 'local:' . $s; }
    function check(): string {
        return marker('x') . '|' . \marker('y');
    }
}
echo \App\check();
"#,
        ["local:x|global:y"]
    };

    namespace_subnamespace_trait_conflict => {
        r#"<?php
namespace A {
    trait HasName { public function name(): string { return 'A'; } }
}
namespace B {
    trait HasName { public function name(): string { return 'B'; } }
}
namespace App {
    use A\HasName as AName;
    class C {
        use AName;
    }
    class D {
        use \B\HasName;
    }
}
echo (new \App\C())->name() . (new \App\D())->name();
"#,
        ["AB"]
    };

    namespace_variable_class_name_with_fqcn => {
        r#"<?php
namespace Domain {
    class Service {
        public function handle(): string { return 'service'; }
    }
}
namespace App {
    $fqcn = 'Domain\\Service';
    $obj = new $fqcn();
    echo $obj->handle();
}
"#,
        ["service"]
    };

    namespace_variable_function_name_from_current_ns => {
        r#"<?php
namespace Runtime {
    function emit(int $n): string { return 'ok:' . $n; }
}
namespace App {
    $f = '\\Runtime\\emit';
    echo $f(3);
}
"#,
        ["ok:3"]
    };

    namespace_alias_collision_prefers_imported_alias => {
        r#"<?php
namespace Local {
    class Logger { public function tag(): string { return 'local'; } }
}
namespace Shared {
    class Logger { public function tag(): string { return 'shared'; } }
}
namespace App {
    use Shared\Logger;
    function pick(): string {
        $a = new Logger();
        $b = new \Local\Logger();
        return $a->tag() . '|' . $b->tag();
    }
}
echo \App\pick();
"#,
        ["shared|local"]
    };

    namespace_root_block_can_create_qualified_and_unqualified_classes => {
        r#"<?php
namespace App {
    class User { public function role(): string { return 'app'; } }
}
namespace {
    function make(): string {
        $user = new App\User();
        $plain = new \App\User();
        return $user->role() . '|' . $plain->role();
    }
    echo make();
}
"#,
        ["app|app"]
    };

    namespace_dynamic_current_namespace_prefix => {
        r#"<?php
namespace Lib {
    class Thing { public function id(): string { return 'thing'; } }
}
namespace App {
    $base = __NAMESPACE__;
    $fqn = $base . '\\Thing';
    $o = new $fqn();
    echo $o->id();
}
"#,
        ["thing"]
    };

    namespace_use_function_alias_chain_and_unprefixed_call => {
        r#"<?php
namespace Tools {
    function normalize(string $s): string { return "n:$s"; }
}
namespace App {
    use function Tools\normalize as norm;
    echo norm('x');
}
"#,
        ["n:x"]
    };

    namespace_function_exists_across_namespaces => {
        r#"<?php
namespace Plugin {
    function enabled(): string { return 'yes'; }
}
namespace App {
    function check(): string {
        return (function_exists('\\Plugin\\enabled') ? 'exists' : 'missing') . '|' . function_exists('Plugin\\enabled');
    }
}
echo \App\check();
"#,
        ["exists|0"]
    };

    namespace_class_exists_with_relative_vs_fqcn => {
        r#"<?php
namespace Core {
    class Engine {}
}
namespace App {
function check(): string {
    $same = class_exists('Engine');
    $fqcn = class_exists('Core\\Engine');
    $withSlash = class_exists('\\Core\\Engine');
        return ($same ? 'same:' : 'same=no:') . ($fqcn ? 'fqcn' : 'nfqcn') . '|' . ($withSlash ? 'slash' : 'nslash');
    }
}
echo \App\check();
"#,
        ["same=no:fqcn|slash"]
    };

    namespace_fully_qualify_current_via_fqn_concat => {
        r#"<?php
namespace Runtime {
    class Worker {
        public function run(): string { return 'ok'; }
    }
}
namespace App {
    $name = __NAMESPACE__ . '\\\\Runtime\\Worker';
    $class = '\\\\' . $name;
    $worker = new $class();
    echo $worker->run();
}
"#,
        ["ok"]
    };

    namespace_relative_to_global_function_lookup_with_use => {
        r#"<?php
function global_marker(string $v): string { return 'g:' . $v; }
namespace App {
    use function global_marker as marker;
    function local(string $v): string { return marker($v); }
    echo local('x');
}
"#,
        ["g:x"]
    };

    namespace_dynamic_function_via_variable_fqcn => {
        r#"<?php
namespace Utils {
    function build(string $s): string { return 'u:' . $s; }
}
namespace App {
    $f = '\\\\Utils\\\\build';
    echo $f('z');
}
"#,
        ["u:z"]
    };

    namespace_trait_prefixed_fully_qualified_use => {
        r#"<?php
namespace Core {
    trait HasToken {
        public function token(): string { return 'tok'; }
    }
}
namespace App {
    class Item {
        use \Core\HasToken;
    }
    echo (new Item())->token();
}
"#,
        ["tok"]
    };

    namespace_aliasing_class_from_subnamespace => {
        r#"<?php
namespace Tools\Data {
    class Payload {
        public function id(): string { return 'payload'; }
    }
}
namespace App {
    use Tools\Data\Payload as DataPayload;
    echo (new DataPayload())->id();
}
"#,
        ["payload"]
    };

    namespace_dynamic_fqcn_variable_reference => {
        r#"<?php
namespace Domain {
    class Worker {
        public function role(): string { return 'worker'; }
    }
}
namespace App {
    $class = 'Domain\\Worker';
    $worker = new $class();
    echo $worker->role();
}
"#,
        ["worker"]
    };

    namespace_nested_resolution_with_current_namespace_prefix => {
        r#"<?php
namespace Infra {
    class Handler {
        public function run(): string { return 'run'; }
    }
}
namespace App\Runtime {
    function make(): string {
        $class = __NAMESPACE__ . '\\\\Handler';
        return (new $class())->run();
    }
    echo make();
}
"#,
        ["run"]
    };

    namespace_variable_function_name_resolves_global_when_qualified => {
        r#"<?php
function marker(string $v): string { return 'global-' . $v; }
namespace App {
    function marker(string $v): string { return 'local-' . $v; }
    $fn = '\\\\marker';
    echo $fn('x');
    echo '|';
    echo function_exists($fn) ? 'exists' : 'missing';
}
"#,
        ["global-x|exists"]
    };

    namespace_function_alias_from_namespace_for_array_walk => {
        r#"<?php
namespace Utility {
    function normalize(string $s): string { return strtoupper($s); }
}
namespace App {
    use function Utility\normalize;
    $items = ['a', 'b'];
    $f = normalize::class;
    $tag = normalize('ok');
    echo $tag;
}
"#,
        ["OK"]
    };

    namespace_invoke_static_call_via_variable_class_name => {
        r#"<?php
namespace Tools {
    class Maker {
        public static function id(string $v): string { return 'tools:' . $v; }
    }
}
namespace App {
    $class = '\\\\Tools\\\\Maker';
    $method = 'id';
    echo $class::$method('x');
}
"#,
        ["tools:x"]
    };

    namespace_class_exists_with_leading_backslash_and_relative => {
        r#"<?php
namespace Core {
    class Engine {}
}
namespace App {
    function check(): string {
        return (
            class_exists('\\Core\\Engine') ? 'leading' : 'noleading'
        ) . '|' . (
            class_exists('Core\\Engine', false) ? 'local' : 'nolocal'
        );
    }
    echo check();
}
"#,
        ["leading|local"]
    };

    namespace_class_reference_from_variable_name => {
        r#"<?php
namespace Domain {
    class Service {
        public function id(): string { return 'service'; }
    }
}
namespace App {
    $class = 'Domain\\Service';
    $svc = new $class();
    echo $svc->id();
}
"#,
        ["service"]
    };

    namespace_call_fully_qualified_name_from_runtime_string => {
        r#"<?php
namespace Shared {
    function helper(int $n): string { return 'n:' . $n; }
}
namespace App {
    $fn = '\\\\Shared\\\\helper';
    echo $fn(7);
}
"#,
        ["n:7"]
    };

    namespace_class_import_does_not_override_function_lookup => {
        r#"<?php
namespace Core {
    function marker(): string { return 'function'; }
    class marker {
        public function value(): string { return 'class'; }
    }
}
namespace App {
    use Core\marker;
    $obj = new marker();
    echo $obj->value();
    echo '|';
    echo \Core\marker();
}
"#,
        ["class|function"]
    };

    namespace_local_and_global_function_resolution => {
        r#"<?php
function marker(): string { return 'global-marker'; }
namespace App {
    function marker(): string { return 'local-marker'; }
    echo marker();
    echo '|';
    echo \marker();
}
"#,
        ["local-marker|global-marker"]
    };

    namespace_trait_alias_and_alias_resolution => {
        r#"<?php
namespace Core {
    trait Marker {
        public function flag(): string { return 'on'; }
    }
}
namespace App {
    use Core\Marker as FlagTrait;
    class Thing {
        use FlagTrait;
    }
    echo (new Thing())->flag();
}
"#,
        ["on"]
    };

    namespace_use_function_alias_chain => {
        r#"<?php
namespace Tools {
    function normalize(string $s): string { return "[$s]"; }
}
namespace App {
    use function Tools\normalize as wrap;
    echo wrap('x');
}
        "#,
        ["[x]"]
    };

    namespace_dynamic_class_lookup_in_block => {
        r#"<?php
namespace Runtime\Loader {
    class Service { public function version(): string { return 'v1'; } }
}
namespace App {
    use Runtime\Loader\Service;
    $name = Service::class;
    $instance = new $name();
    echo $instance->version();
}
"#,
        ["v1"]
    };

    namespace_fully_qualified_function_isolation => {
        r#"<?php
namespace Local {
    function format(string $v): string { return "local:$v"; }
}
namespace {
    function format(string $v): string { return "global:$v"; }
    echo \Local\format('a');
    echo '|';
    echo format('b');
}
"#,
        ["local:a|global:b"]
    };

    namespace_backslash_fqcn_with_class_exists_check => {
        r#"<?php
namespace Services {
    class Worker { public static function ok(): string { return 'yes'; } }
}
namespace App {
    echo class_exists('Services\\\\Worker') ? 'yes' : 'no';
}
"#,
        ["yes"]
    };

    namespace_backslash_global_function_call_from_namespaced => {
        r#"<?php
namespace App {
    echo \strlen('hey');
}
"#,
        ["3"]
    };

    namespace_group_use_with_class_and_function => {
        r#"<?php
namespace Core {
    class Box { public function name(): string { return 'box'; } }
    function helper(string $name): string { return "h:$name"; }
}
namespace App {
    use Core\{Box, function helper};
    $b = new Box();
    echo $b->name();
    echo '|';
    echo helper('x');
}
"#,
        ["box|h:x"]
    };

    namespace_imported_trait_method_visibility => {
        r#"<?php
namespace Core\Traits {
    trait Logger {
        protected function tag(): string { return 'tag'; }
    }
}
namespace App {
    class Thing {
        use Core\Traits\Logger {
            tag as public;
        }
    }
    echo (new Thing())->tag();
}
"#,
        ["tag"]
    };

    namespace_nested_alias_stack => {
        r#"<?php
namespace Infra {
    class Handler { public function id(): string { return 'handler'; } }
}
namespace App\Services {
    use Infra\Handler as InfraHandler;
    class Facade {
        public function make(): string {
            return (new InfraHandler())->id();
        }
    }
}
namespace App {
    use App\Services\Facade;
    echo (new Facade())->make();
}
"#,
        ["handler"]
    };

    namespace_function_alias_chain_with_local_shadow => {
        r#"<?php
namespace Utils {
    function ping(string $value): string { return "u:$value"; }
}
namespace App {
    function ping(string $value): string { return "a:$value"; }
    use function Utils\ping as util_ping;
    function run(string $value): string {
        return ping($value) . '|' . util_ping($value);
    }
    echo run('x');
}
"#,
        ["a:x|u:x"]
    };

    namespace_constant_alias_and_unqualified_use => {
        r#"<?php
namespace Config {
    const MODE = 'live';
}
namespace App {
    use const Config\MODE as CurrentMode;
    function status(): string {
        return "mode:" . CurrentMode;
    }
    echo status();
}
"#,
        ["mode:live"]
    };

    namespace_class_alias_between_namespaces => {
        r#"<?php
namespace Domain {
    class Repo { public function type(): string { return 'repo'; } }
}
namespace App {
    use Domain\Repo as Repository;
    echo (new Repository())->type();
}
"#,
        ["repo"]
    };

    namespace_qualified_call_via_fully_qualified_function_string => {
        r#"<?php
namespace Library {
    function label(string $value): string { return "label:$value"; }
}
namespace App {
    $f = '\\\\Library\\\\label';
    echo $f('abc');
}
"#,
        ["label:abc"]
    };

    namespace_local_namespace_prefers_local_class_without_use => {
        r#"<?php
namespace Shared {
    class Logger { public function channel(): string { return 'shared'; } }
}
namespace App {
    class Logger { public function channel(): string { return 'app'; } }
    $local = new Logger();
    $global = new \Shared\Logger();
    echo $local->channel() . '|' . $global->channel();
}
"#,
        ["app|shared"]
    };

    namespace_function_call_resolution_with_use_as_alias_chain => {
        r#"<?php
namespace Runtime {
    function normalize(string $value): string { return strtoupper($value); }
}
namespace App {
    use function Runtime\normalize as up;
    echo up('ok');
}
"#,
        ["OK"]
    };

    namespace_class_exists_for_leading_backslash_current_ns => {
        r#"<?php
namespace Tooling {
    class Engine {}
}
namespace App {
    function check(): string {
        $local = class_exists('Tooling\\Engine', false);
        $global = class_exists('\\Tooling\\Engine');
        return ($local ? 'local' : 'nolocal') . '|' . ($global ? 'global' : 'noglobal');
    }
    echo check();
}
"#,
        ["local|global"]
    };

    namespace_dynamic_fqcn_constructor_in_rooted_namespace => {
        r#"<?php
namespace Workers {
    class Agent {
        public function job(): string { return 'done'; }
    }
}
namespace App {
    $fqn = '\\\\Workers\\\\Agent';
    $agent = new $fqn();
    echo $agent->job();
}
"#,
        ["done"]
    };

    namespace_fqcn_variable_static_call => {
        r#"<?php
namespace Services {
    class Factory {
        public static function make(int $n): int { return $n + 100; }
    }
}
namespace App {
    $class = '\\Services\\Factory';
    $method = 'make';
    echo $class::$method(9);
}
"#,
        ["109"]
    };

    namespace_function_name_from_namespace_context => {
        r#"<?php
namespace Tools {
    function label(string $s): string { return 'tool:' . $s; }
}
namespace App {
    $fn = '\\Tools\\\\label';
    echo $fn('x');
}
"#,
        ["tool:x"]
    };

    namespace_dynamic_function_from_variable_namespace_string => {
        r#"<?php
namespace Util {
    function stamp(string $s): string { return 'stamp:' . $s; }
}
namespace App {
    $ns = '\\\\Util';
    $fn = $ns . '\\\\stamp';
    echo $fn('done');
}
"#,
        ["stamp:done"]
    };

    namespace_alias_chain_with_global_fallback => {
        r#"<?php
namespace {
    function marker(string $v): string { return 'global:' . $v; }
}
namespace Package {
    function marker(string $v): string { return 'local:' . $v; }
}
namespace App {
    use function Package\\marker as local_marker;
    echo local_marker('x') . '|' . \Package\marker('y');
}
"#,
        ["local:x|local:y"]
    };

    namespace_group_use_mixed_and_nested_aliases => {
        r#"<?php
namespace Core {
    function lower(string $s): string { return strtolower($s); }
    const CODE = 'ok';
}
namespace App {
    use function Core\lower as down;
    use const Core\CODE;
    echo down(CODE);
}
"#,
        ["ok"]
    };
}

#[test]
fn namespace_current_constant_from_nested_block() {
    compile_ok(
        r#"<?php
namespace Shop\Models {
    class Product {
        public function scope(): string { return __NAMESPACE__; }
    }
}
namespace App {
    $obj = new \Shop\Models\Product();
    echo $obj->scope();
}
"#,
    );
}

#[test]
fn namespace_fully_qualified_calls_from_root() {
    compile_ok(
        r#"<?php
namespace App {
    function tag(string $v): string { return "app:$v"; }
}
function tag(string $v): string { return "global:$v"; }
echo \App\tag('x') === 'app:x' ? 'app' : 'no';
echo '|' . tag('y') === 'global:y' ? 'global' : 'noglobal';
"#,
    );
}

#[test]
fn namespace_use_import_in_anonymous_function_scope() {
    compile_ok(
        r#"<?php
namespace Lib {
    function norm(string $v): string { return "[$v]"; }
}
namespace App {
    use function Lib\norm;
    $render = function() use () {
        return norm('x');
    };
    echo $render();
}
"#,
    );
}

#[test]
fn namespace_dynamic_class_name_with_braces() {
    compile_ok(
        r#"<?php
namespace Domain {
    class Service { public function label(): string { return 'service'; } }
}
namespace App {
    $c = 'Service';
    $fqcn = "\\Domain\\$c";
    $obj = new $fqcn();
    echo $obj->label();
}
"#,
    );
}
