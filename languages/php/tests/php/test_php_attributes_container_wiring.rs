//! Attributes driving behaviour: dependency-injection wiring, listener
//! priority ordering, and attribute instances used as live objects. This is
//! the half of the framework contract where `newInstance()` produces something
//! the container then *calls*, not just reads.
//!
//! Expected values generated from PHP 8.4.11.

crate::php_cases! {
    constructor_parameter_attributes_drive_container_resolution => {
        r#"<?php
#[Attribute(Attribute::TARGET_PARAMETER)]
class Autowire {
    public function __construct(public string $service) {}
}
class Mailer {
    public function send(): string { return 'sent'; }
}
class Notifier {
    public function __construct(
        #[Autowire('mailer')] public Mailer $m,
        #[Autowire('logger')] public $log
    ) {}
}
$container = ['mailer' => new Mailer(), 'logger' => 'LOG'];
$args = [];
foreach ((new ReflectionClass(Notifier::class))->getConstructor()->getParameters() as $p) {
    $a = $p->getAttributes(Autowire::class);
    $args[] = $a ? $container[$a[0]->newInstance()->service] : null;
}
$n = new Notifier(...$args);
echo $n->m->send() . '/' . $n->log;
"#,
        ["sent/LOG"]
    };

    get_target_reports_where_the_attribute_was_applied => {
        r#"<?php
#[Attribute(Attribute::TARGET_CLASS | Attribute::TARGET_METHOD)]
class Tag {}
#[Tag]
class Klass {
    #[Tag]
    public function m() {}
}
$ct = (new ReflectionClass(Klass::class))->getAttributes(Tag::class)[0]->getTarget();
$mt = (new ReflectionMethod(Klass::class, 'm'))->getAttributes(Tag::class)[0]->getTarget();
echo $ct . ',' . $mt . ',' . Attribute::TARGET_CLASS . ',' . Attribute::TARGET_METHOD;
"#,
        ["1,4,1,4"]
    };

    listener_priority_from_attributes_orders_the_dispatch_chain => {
        r#"<?php
#[Attribute]
class AsListener {
    public function __construct(public string $event, public int $priority = 0) {}
}
#[AsListener('kernel.request', priority: 10)]
class Auth {}
#[AsListener('kernel.request', priority: 100)]
class Firewall {}
#[AsListener('kernel.request')]
class Router {}
$ls = [];
foreach ([Auth::class, Firewall::class, Router::class] as $c) {
    $a = (new ReflectionClass($c))->getAttributes(AsListener::class)[0]->newInstance();
    $ls[] = [$c, $a->priority];
}
usort($ls, fn($x, $y) => $y[1] <=> $x[1]);
echo implode('>', array_map(fn($l) => $l[0] . '(' . $l[1] . ')', $ls));
"#,
        ["Firewall(100)>Auth(10)>Router(0)"]
    };

    attribute_instance_method_is_callable_after_new_instance => {
        r#"<?php
#[Attribute]
class Transform {
    public function __construct(private string $prefix) {}
    public function apply(string $v): string { return $this->prefix . $v; }
}
class Field {
    #[Transform('pre-')]
    public $name;
}
$t = (new ReflectionProperty(Field::class, 'name'))->getAttributes(Transform::class)[0]->newInstance();
echo $t->apply('value');
"#,
        ["pre-value"]
    };

    is_instanceof_scan_collects_validators_by_shared_interface => {
        r#"<?php
interface Constraint {
    public function check(int $v): bool;
}
#[Attribute]
class Min implements Constraint {
    public function __construct(private int $n) {}
    public function check(int $v): bool { return $v >= $this->n; }
}
#[Attribute]
class Max implements Constraint {
    public function __construct(private int $n) {}
    public function check(int $v): bool { return $v <= $this->n; }
}
class Form {
    #[Min(5)]
    #[Max(10)]
    public int $age = 0;
}
$rp = new ReflectionProperty(Form::class, 'age');
$ok = [];
foreach ($rp->getAttributes(Constraint::class, ReflectionAttribute::IS_INSTANCEOF) as $a) {
    $ok[] = $a->newInstance()->check(7) ? 'y' : 'n';
}
echo count($ok) . ':' . implode('', $ok);
"#,
        ["2:yy"]
    };
}
