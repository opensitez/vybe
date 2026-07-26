//! Attribute *discovery* — the scan-and-build pattern Symfony and Laravel use
//! for routing. Unlike the unit-level tests in `test_attributes.rs`, nothing
//! here indexes a known reflection target directly: every case walks
//! `getMethods()` / `getProperties()` and builds a table from what it finds.
//!
//! Expected values generated from PHP 8.4.11.

crate::php_cases! {
    route_table_built_by_scanning_controller_methods => {
        r#"<?php
#[Attribute(Attribute::TARGET_METHOD)]
class Route {
    public function __construct(public string $path, public string $method = 'GET') {}
}
class UserController {
    #[Route('/users')]
    public function index(): string { return 'index'; }
    #[Route('/users/{id}')]
    public function show(): string { return 'show'; }
    public function helper(): string { return 'helper'; }
    #[Route('/users', method: 'POST')]
    public function create(): string { return 'create'; }
}
$routes = [];
foreach ((new ReflectionClass(UserController::class))->getMethods() as $m) {
    foreach ($m->getAttributes(Route::class) as $a) {
        $r = $a->newInstance();
        $routes[] = $r->method . ' ' . $r->path . ' -> ' . $m->getName();
    }
}
echo implode('|', $routes);
"#,
        ["GET /users -> index|GET /users/{id} -> show|POST /users -> create"]
    };

    discovery_skips_methods_without_the_attribute => {
        r#"<?php
#[Attribute]
class Route {
    public function __construct(public string $path) {}
}
class C {
    #[Route('/a')]
    public function a() {}
    public function b() {}
    #[Route('/c')]
    public function c() {}
}
$n = 0;
$paths = [];
foreach ((new ReflectionClass(C::class))->getMethods() as $m) {
    $n++;
    if ($a = $m->getAttributes(Route::class)) {
        $paths[] = $a[0]->newInstance()->path;
    }
}
echo $n . ':' . implode(',', $paths);
"#,
        ["3:/a,/c"]
    };

    discovery_collects_repeated_attributes_per_method => {
        r#"<?php
#[Attribute(Attribute::IS_REPEATABLE | Attribute::TARGET_METHOD)]
class Verb {
    public function __construct(public string $name) {}
}
class Api {
    #[Verb('GET')]
    #[Verb('HEAD')]
    public function read(): string { return 'r'; }
}
$verbs = [];
foreach ((new ReflectionClass(Api::class))->getMethods() as $m) {
    foreach ($m->getAttributes(Verb::class) as $a) {
        $verbs[] = $a->newInstance()->name;
    }
}
echo implode(',', $verbs);
"#,
        ["GET,HEAD"]
    };

    class_level_prefix_combines_with_method_routes_across_classes => {
        r#"<?php
#[Attribute(Attribute::TARGET_CLASS)]
class Prefix {
    public function __construct(public string $base) {}
}
#[Attribute(Attribute::TARGET_METHOD)]
class Route {
    public function __construct(public string $path) {}
}
#[Prefix('/api/users')]
class UserApi {
    #[Route('/list')]
    public function list() {}
}
#[Prefix('/api/posts')]
class PostApi {
    #[Route('/list')]
    public function list() {}
    #[Route('/new')]
    public function new_() {}
}
$out = [];
foreach ([UserApi::class, PostApi::class] as $c) {
    $rc = new ReflectionClass($c);
    $pre = $rc->getAttributes(Prefix::class)[0]->newInstance()->base;
    foreach ($rc->getMethods() as $m) {
        foreach ($m->getAttributes(Route::class) as $a) {
            $out[] = $pre . $a->newInstance()->path;
        }
    }
}
echo implode(' ', $out);
"#,
        ["/api/users/list /api/posts/list /api/posts/new"]
    };

    property_scan_builds_orm_column_map => {
        r#"<?php
#[Attribute]
class Column {
    public function __construct(public string $type, public bool $nullable = false) {}
}
#[Attribute]
class Id {}
class User {
    #[Id]
    #[Column('integer')]
    public int $id = 0;
    #[Column('string', nullable: true)]
    public ?string $email = null;
    public string $ignored = '';
}
$map = [];
foreach ((new ReflectionClass(User::class))->getProperties() as $p) {
    $cols = $p->getAttributes(Column::class);
    if (!$cols) continue;
    $c = $cols[0]->newInstance();
    $pk = $p->getAttributes(Id::class) ? '*' : '';
    $map[] = $pk . $p->getName() . ':' . $c->type . ($c->nullable ? '?' : '');
}
echo implode('|', $map);
"#,
        ["*id:integer|email:string?"]
    };
}
